using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelReadDiagnosticsTests
{
    [Fact]
    public void MissingRequiredHeaderProvidesStructuredHeaderDiagnostic()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<Product>().Column(product => product.Name).Column(product => product.Quantity).Build();

        try
        {
            CreateWorkbook(path, ["Name"], ["Shirt"]);
            var error = Assert.Single(Excel.Read(path, schema).Errors);

            Assert.Equal(ExcelReadErrorCode.MissingHeader, error.Code);
            Assert.Equal(ExcelReadErrorKind.MissingHeader, error.Kind);
            Assert.Equal(1U, error.Row);
            Assert.Null(error.Column);
            Assert.Equal("Quantity", error.Header);
            Assert.Equal(nameof(Product.Quantity), error.PropertyName);
            Assert.Null(error.RawValue);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void OptionalMissingHeaderProducesNoDiagnosticButInvalidOptionalValueDoes()
    {
        var missingPath = TemporaryPath();
        var invalidPath = TemporaryPath();
        var schema = Excel.Schema<Product>()
            .Column(product => product.Name)
            .Column(product => product.OptionalQuantity, column => column.Optional())
            .Build();

        try
        {
            CreateWorkbook(missingPath, ["Name"], ["Shirt"]);
            CreateWorkbook(invalidPath, ["Name", "OptionalQuantity"], ["Shirt", "bad"]);

            Assert.True(Excel.Read(missingPath, schema).IsValid);
            var error = Assert.Single(Excel.Read(invalidPath, schema).Errors);
            Assert.Equal(ExcelReadErrorKind.Conversion, error.Kind);
            Assert.Equal(nameof(Product.OptionalQuantity), error.PropertyName);
            Assert.Equal("bad", error.RawValue);
        }
        finally
        {
            Delete(missingPath);
            Delete(invalidPath);
        }
    }

    [Fact]
    public void ConversionAndValidationDiagnosticsExposeCoordinatesPropertiesAndRawValues()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<Product>()
            .Column(product => product.Name)
            .Column(product => product.Quantity, column => column.Range(1, 10).Validate(value => value % 2 == 0, "Quantity must be even"))
            .Build();

        try
        {
            CreateWorkbook(path, ["Name", "Quantity"], ["Broken", "abc"], ["Odd", "11"]);
            var result = Excel.Read(path, schema);
            var conversion = Assert.Single(result.Errors, error => error.Kind == ExcelReadErrorKind.Conversion);
            var validations = result.Errors.Where(error => error.Kind == ExcelReadErrorKind.Validation).ToArray();

            Assert.Equal(2U, conversion.Row);
            Assert.Equal(2, conversion.Column);
            Assert.Equal("B2", conversion.CellReference);
            Assert.Equal("Quantity", conversion.Header);
            Assert.Equal(nameof(Product.Quantity), conversion.PropertyName);
            Assert.Equal("abc", conversion.RawValue);
            Assert.Equal(typeof(int), conversion.ExpectedType);
            Assert.Equal(2, validations.Length);
            Assert.All(validations, error =>
            {
                Assert.Equal(3U, error.Row);
                Assert.Equal("11", error.RawValue);
                Assert.Equal(nameof(Product.Quantity), error.PropertyName);
            });
            Assert.Contains(validations, error => error.Message.Contains("between 1 and 10", StringComparison.Ordinal));
            Assert.Contains(validations, error => error.Message == "Quantity must be even");
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ConverterFailureIsAConversionDiagnosticAndConverterExceptionsStillPropagate()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<ConvertedProduct>()
            .Column(product => product.Status, column => column.ConvertWith(StatusConverter.Instance))
            .Build();

        try
        {
            CreateWorkbook(path, ["Status"], ["unknown"]);
            var error = Assert.Single(Excel.Read(path, schema).Errors);

            Assert.Equal(ExcelReadErrorKind.Conversion, error.Kind);
            Assert.Equal(nameof(ConvertedProduct.Status), error.PropertyName);
            Assert.Equal("unknown", error.RawValue);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void DeclarativeListAndDateValidationProduceValidationDiagnostics()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<ValidationProduct>()
            .Column(product => product.Status, column => column.AllowedValues("Active", "Inactive"))
            .Column(product => product.AvailableOn, column => column.DateBetween(new DateTime(2026, 1, 1), new DateTime(2026, 12, 31)))
            .Build();

        try
        {
            CreateWorkbook(path, ["Status", "AvailableOn"], ["Archived", "2027-01-01"]);
            var errors = Excel.Read(path, schema).Errors;

            Assert.Equal(2, errors.Count);
            Assert.All(errors, error => Assert.Equal(ExcelReadErrorKind.Validation, error.Kind));
            Assert.Contains(errors, error => error.PropertyName == nameof(ValidationProduct.Status)
                && Equals("Archived", error.RawValue)
                && error.Message.Contains("one of", StringComparison.Ordinal));
            Assert.Contains(errors, error => error.PropertyName == nameof(ValidationProduct.AvailableOn)
                && Equals("2027-01-01", error.RawValue)
                && error.Message.Contains("between", StringComparison.Ordinal));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void IncludePolicyReturnsPartiallyMaterializedInvalidRowsWhileDefaultSkipsThem()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<Product>().Column(product => product.Name).Column(product => product.Quantity).Build();

        try
        {
            CreateWorkbook(path, ["Name", "Quantity"], ["Valid", "2"], ["Broken", "abc"]);

            var skipped = Excel.Read(path, schema);
            var included = Excel.Read(path, schema, new ExcelReadOptions(ExcelInvalidRowPolicy.Include));

            Assert.Equal("Valid", Assert.Single(skipped.Rows).Name);
            Assert.Equal(2, included.Rows.Count);
            Assert.Equal("Broken", included.Rows[1].Name);
            Assert.Equal(0, included.Rows[1].Quantity);
            Assert.False(included.IsValid);
            Assert.Single(included.Errors);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void DisabledOverlayColumnsDoNotCreateErrorsAndRuntimeHeadersAreReported()
    {
        var disabledPath = TemporaryPath();
        var runtimeHeaderPath = TemporaryPath();
        var schema = Excel.Schema<Product>()
            .Column(product => product.Name)
            .Column(product => product.Quantity)
            .Build();
        var disabled = Excel.Overlay<Product>(overlay => overlay.Include(product => product.Quantity, false));
        var translated = Excel.Overlay<Product>(overlay => overlay.Header(product => product.Quantity, "Quantità"));

        try
        {
            CreateWorkbook(disabledPath, ["Name", "Quantity"], ["Shirt", "bad"]);
            CreateWorkbook(runtimeHeaderPath, ["Name", "Quantità"], ["Shirt", "bad"]);

            var disabledResult = Excel.Read(disabledPath, schema, disabled, new ExcelReadOptions(ExcelInvalidRowPolicy.Include));
            var runtimeError = Assert.Single(Excel.Read(runtimeHeaderPath, schema, translated).Errors);

            Assert.True(disabledResult.IsValid);
            Assert.Equal("Shirt", Assert.Single(disabledResult.Rows).Name);
            Assert.Equal("Quantità", runtimeError.Header);
            Assert.Equal(nameof(Product.Quantity), runtimeError.PropertyName);
        }
        finally
        {
            Delete(disabledPath);
            Delete(runtimeHeaderPath);
        }
    }

    [Fact]
    public void PositionalMappingReportsThePhysicalExcelColumn()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<Product>()
            .Column(product => product.Name, column => column.Optional())
            .Column(product => product.Quantity, column => column.MapFromColumn(3))
            .Build();
        var overlay = Excel.Overlay<Product>(runtime => runtime.Include(product => product.Name, false));

        try
        {
            CreateWorkbook(path, ["Ignored", "Also ignored", "Quantity"], ["x", "y", "bad"]);
            var error = Assert.Single(Excel.Read(path, schema, overlay).Errors);

            Assert.Equal(3, error.Column);
            Assert.Equal("C2", error.CellReference);
            Assert.Equal(nameof(Product.Quantity), error.PropertyName);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void WorkbookReadsKeepPoliciesAndDiagnosticsIndependent()
    {
        var path = TemporaryPath();
        var customerSchema = Excel.Schema<Product>().SheetName("Customers").Column(product => product.Name).Column(product => product.Quantity).Build();
        var orderSchema = Excel.Schema<Order>().SheetName("Orders").Column(order => order.Id).Build();

        try
        {
            Excel.Workbook()
                .AddSheet([new Product { Name = "Ada", Quantity = 1 }], customerSchema)
                .AddSheet([new Order { Id = 7 }], orderSchema)
                .Write(path);
            ReplaceDataCell(path, "Customers", 1, "bad");

            using var workbook = Excel.Open(path);
            var customers = workbook.Read(customerSchema, new ExcelReadOptions(ExcelInvalidRowPolicy.Include));
            var orders = workbook.ReadAt(1, orderSchema, new ExcelReadOptions(ExcelInvalidRowPolicy.Skip));

            Assert.False(customers.IsValid);
            Assert.Equal("Customers", Assert.Single(customers.Errors).SheetName);
            Assert.Equal("Ada", Assert.Single(customers.Rows).Name);
            Assert.True(orders.IsValid);
            Assert.Equal(7, Assert.Single(orders.Rows).Id);
        }
        finally
        {
            Delete(path);
        }
    }

    private static void CreateWorkbook(string path, string[] headers, params string[][] rows)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var data = new SheetData(CreateRow(headers));
        foreach (var row in rows)
        {
            data.AppendChild(CreateRow(row));
        }

        worksheetPart.Worksheet = new Worksheet(data);
        var sheets = workbookPart.Workbook!.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1U, Name = "Sheet1" });
        worksheetPart.Worksheet!.Save();
        workbookPart.Workbook!.Save();
    }

    private static Row CreateRow(IEnumerable<string> values) => new(values.Select(value => new Cell
    {
        DataType = CellValues.InlineString,
        InlineString = new InlineString(new Text(value)),
    }));

    private static void ReplaceDataCell(string path, string sheetName, int cellIndex, string value)
    {
        using var document = SpreadsheetDocument.Open(path, true);
        var sheet = document.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().Single(item => item.Name!.Value == sheetName);
        var worksheet = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id!);
        var cell = worksheet.Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().Skip(1).Single().Elements<Cell>().ElementAt(cellIndex);
        cell.RemoveAllChildren();
        cell.DataType = CellValues.InlineString;
        cell.InlineString = new InlineString(new Text(value));
        worksheet.Worksheet!.Save();
    }

    private static string TemporaryPath() => Path.Combine(Path.GetTempPath(), $"cellsharp-read-diagnostics-{Guid.NewGuid():N}.xlsx");

    private static void Delete(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }

    private sealed class Product
    {
        public string? Name { get; set; }
        public int Quantity { get; set; }
        public int? OptionalQuantity { get; set; }
    }

    private sealed class Order
    {
        public int Id { get; set; }
    }

    private sealed class ValidationProduct
    {
        public string? Status { get; set; }
        public DateTime AvailableOn { get; set; }
    }

    private sealed class ConvertedProduct
    {
        public Status? Status { get; set; }
    }

    private sealed class Status
    {
        internal Status(string value) => Value = value;
        internal string Value { get; }
    }

    private sealed class StatusConverter : IExcelValueConverter<Status?, string>
    {
        internal static StatusConverter Instance { get; } = new();

        public string Write(Status? value) => value?.Value ?? string.Empty;

        public bool TryRead(string value, out Status? converted)
        {
            converted = null;
            return false;
        }
    }
}
