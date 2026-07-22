using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelWorkbookTests
{
    [Fact]
    public void WorkbookWritesHeterogeneousSheetsInConfiguredOrderWithSharedValidationLookup()
    {
        var path = TemporaryPath();
        var customerSchema = CustomerSchema("Customers");
        var orderSchema = OrderSchema("Orders");

        try
        {
            Excel.Workbook()
                .AddSheet(
                    [new Customer { Name = "Ada", Status = Status.Active, Category = "Retail" }],
                    customerSchema,
                    options => options.Theme(ExcelTheme.Modern).FreezeHeaderRow())
                .AddSheet(
                    [new Order { Id = 7, Amount = 123.45M, Category = "Retail" }],
                    orderSchema,
                    options => options.Theme(ExcelTheme.Classic).AutoFitColumns().AlternatingRows())
                .Write(path);

            using var document = SpreadsheetDocument.Open(path, false);
            var sheets = document.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().ToArray();
            var validationSheet = Assert.Single(sheets, sheet => sheet.State?.Value == SheetStateValues.Hidden);
            var customerSheet = sheets[0];
            var orderSheet = sheets[1];
            var customerWorksheet = Worksheet(document, customerSheet);
            var orderWorksheet = Worksheet(document, orderSheet);
            var orderAmount = orderWorksheet.GetFirstChild<SheetData>()!.Elements<Row>().Skip(1).Single().Elements<Cell>().ElementAt(1);

            Assert.Equal(["Customers", "Orders", "_CellSharpValidation"], sheets.Select(sheet => sheet.Name!.Value));
            Assert.Equal(SheetStateValues.Hidden, validationSheet.State!.Value);
            Assert.NotNull(customerWorksheet.GetFirstChild<DataValidations>());
            Assert.NotNull(orderWorksheet.GetFirstChild<DataValidations>());
            Assert.NotNull(orderWorksheet.GetFirstChild<Columns>());
            Assert.NotEqual(0U, orderAmount.StyleIndex!.Value);
            Assert.All(Worksheet(document, validationSheet).GetFirstChild<SheetData>()!.Elements<Row>(), row => Assert.Single(row.Elements<Cell>()));
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void WorkbookSupportsThreeSheetsAndRejectsCaseInsensitiveDuplicateNames()
    {
        var path = TemporaryPath();
        var first = CustomerSchema("First");
        var second = CustomerSchema("Second");
        var third = CustomerSchema("Third");

        try
        {
            Excel.Workbook()
                .AddSheet([new Customer { Name = "A" }], first)
                .AddSheet([new Customer { Name = "B" }], second)
                .AddSheet([new Customer { Name = "C" }], third)
                .Write(path);

            using var document = SpreadsheetDocument.Open(path, false);
            Assert.Equal(["First", "Second", "Third", "_CellSharpValidation"], document.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().Select(sheet => sheet.Name!.Value));

            Assert.Throws<ArgumentException>(() => Excel.Workbook()
                .AddSheet(Array.Empty<Customer>(), CustomerSchema("People"))
                .AddSheet(Array.Empty<Customer>(), CustomerSchema("people")));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void WorkbookKeepsUserSheetNamesWhenValidationLookupWouldCollide()
    {
        var path = TemporaryPath();

        try
        {
            Excel.Workbook()
                .AddSheet([new Customer { Name = "Ada", Category = "Retail" }], CustomerSchema("_CellSharpValidation"))
                .AddSheet([new Order { Id = 7, Amount = 1M, Category = "Wholesale" }], OrderSchema("Orders"))
                .Write(path);

            using var document = SpreadsheetDocument.Open(path, false);
            var sheets = document.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().ToArray();

            Assert.Equal(["_CellSharpValidation", "Orders", "_CellSharpValidation2"], sheets.Select(sheet => sheet.Name!.Value));
            Assert.Null(sheets[0].State);
            Assert.Equal(SheetStateValues.Hidden, sheets[2].State!.Value);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void WorkbookCreatesMultiSheetTemplatesWithIndependentPresentationAndValidation()
    {
        var path = TemporaryPath();
        var customerSchema = CustomerSchema("Customers");
        var orderSchema = OrderSchema("Orders");

        try
        {
            Excel.Workbook()
                .AddTemplateSheet(customerSchema, options => options.Theme(ExcelTheme.Modern).FreezeHeaderRow())
                .AddTemplateSheet(orderSchema, options => options.AutoFitColumns())
                .CreateTemplate(path);

            using var document = SpreadsheetDocument.Open(path, false);
            var sheets = document.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().ToArray();

            Assert.Equal(["Customers", "Orders", "_CellSharpValidation"], sheets.Select(sheet => sheet.Name!.Value));
            Assert.Single(Worksheet(document, sheets[0]).GetFirstChild<SheetData>()!.Elements<Row>());
            Assert.Single(Worksheet(document, sheets[1]).GetFirstChild<SheetData>()!.Elements<Row>());
            Assert.NotNull(Worksheet(document, sheets[0]).GetFirstChild<DataValidations>());
            Assert.NotNull(Worksheet(document, sheets[1]).GetFirstChild<DataValidations>());
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void OpenReadsMultipleTypedSheetsWithoutReopeningTheWorkbook()
    {
        var path = TemporaryPath();
        var customerSchema = CustomerSchema("Customers");
        var orderSchema = OrderSchema("Orders");

        try
        {
            Excel.Workbook()
                .AddSheet([new Customer { Name = "Ada", Status = Status.Active, Category = "Retail" }], customerSchema)
                .AddSheet([new Order { Id = 7, Amount = 123.45M, Category = "Wholesale" }], orderSchema)
                .Write(path);

            using var workbook = Excel.Open(path);
            var customers = workbook.Read(customerSchema);
            var orders = workbook.Read(orderSchema);

            Assert.Equal(["Customers", "Orders"], workbook.WorksheetNames);
            Assert.Equal("Ada", Assert.Single(customers.Rows).Name);
            Assert.Equal(Status.Active, customers.Rows[0].Status);
            Assert.Equal(7, Assert.Single(orders.Rows).Id);
            Assert.Equal(123.45M, orders.Rows[0].Amount);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void OpenSupportsZeroBasedPublicSheetIndexesAndRejectsIncompatibleSelection()
    {
        var path = TemporaryPath();
        var customerSchema = CustomerSchema("Customers");
        var orderSchema = OrderSchema("Orders");

        try
        {
            Excel.Workbook()
                .AddSheet([new Customer { Name = "Ada" }], customerSchema)
                .AddSheet([new Order { Id = 7, Amount = 1M }], orderSchema)
                .Write(path);

            using var workbook = Excel.Open(path);

            Assert.Equal("Ada", Assert.Single(workbook.ReadAt(0, customerSchema).Rows).Name);
            Assert.Equal(7, Assert.Single(workbook.ReadAt(1, orderSchema).Rows).Id);
            Assert.Throws<InvalidOperationException>(() => workbook.ReadAt(0, orderSchema));
            Assert.Throws<ArgumentOutOfRangeException>(() => workbook.ReadAt(2, customerSchema));
            Assert.Throws<InvalidOperationException>(() => workbook.Read(CustomerSchema("Missing")));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void OpenReadsExplicitlyNamedHiddenUserSheetsAndRejectsInternalSheets()
    {
        var hiddenPath = TemporaryPath();
        var validationPath = TemporaryPath();
        var hiddenSchema = Excel.Schema<Customer>()
            .SheetName("Archive")
            .Column(customer => customer.Name)
            .Build();
        var customerSchema = CustomerSchema("Customers");

        try
        {
            CreateHiddenWorkbook(hiddenPath, "Archive", "Name", "Ada");
            Excel.Workbook()
                .AddSheet([new Customer { Name = "Ada", Category = "Retail" }], customerSchema)
                .Write(validationPath);

            using var hidden = Excel.Open(hiddenPath);
            using var validation = Excel.Open(validationPath);

            Assert.Equal("Ada", Assert.Single(hidden.Read(hiddenSchema).Rows).Name);
            Assert.Throws<InvalidOperationException>(() => validation.Read(CustomerSchema("_CellSharpValidation")));
        }
        finally
        {
            Delete(hiddenPath);
            Delete(validationPath);
        }
    }

    [Fact]
    public void InvalidDataInOneSheetDoesNotInterfereWithAnotherTypedRead()
    {
        var path = TemporaryPath();
        var customerSchema = CustomerSchema("Customers");
        var orderSchema = OrderSchema("Orders");

        try
        {
            Excel.Workbook()
                .AddSheet([new Customer { Name = "Ada" }], customerSchema)
                .AddSheet([new Order { Id = 7, Amount = 1M }], orderSchema)
                .Write(path);
            ReplaceFirstDataCellWithText(path, "Orders", "invalid");

            using var workbook = Excel.Open(path);
            var customers = workbook.Read(customerSchema);
            var orders = workbook.Read(orderSchema);

            Assert.True(customers.IsValid);
            Assert.Equal("Ada", Assert.Single(customers.Rows).Name);
            Assert.Empty(orders.Rows);
            Assert.Equal(ExcelReadErrorCode.InvalidValue, Assert.Single(orders.Errors).Code);
        }
        finally
        {
            Delete(path);
        }
    }

    private static ExcelSchema<Customer> CustomerSchema(string sheetName) => Excel.Schema<Customer>()
        .SheetName(sheetName)
        .Column(customer => customer.Name)
        .Column(customer => customer.Status, column => column.ConvertWith(StatusConverter.Instance))
        .Column(customer => customer.Category, column => column.Optional().AllowedValues("Retail", "Wholesale"))
        .Build();

    private static ExcelSchema<Order> OrderSchema(string sheetName) => Excel.Schema<Order>()
        .SheetName(sheetName)
        .Column(order => order.Id)
        .Column(order => order.Amount, column => column.Format("#,##0.00").Width(16D))
        .Column(order => order.Category, column => column.Optional().AllowedValues("Retail", "Wholesale"))
        .Build();

    private static Worksheet Worksheet(SpreadsheetDocument document, Sheet sheet) => (document.WorkbookPart!.GetPartById(sheet.Id!) as WorksheetPart)!
        .Worksheet!;

    private static void CreateHiddenWorkbook(string path, string sheetName, string header, string value)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var data = new SheetData(
            new Row(new Cell
            {
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new Text(header)),
            }),
            new Row(new Cell
            {
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new Text(value)),
            }));
        worksheetPart.Worksheet = new Worksheet(data);
        var sheets = workbookPart.Workbook!.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1U,
            Name = sheetName,
            State = SheetStateValues.Hidden,
        });
        worksheetPart.Worksheet!.Save();
        workbookPart.Workbook!.Save();
    }

    private static void ReplaceFirstDataCellWithText(string path, string sheetName, string value)
    {
        using var document = SpreadsheetDocument.Open(path, true);
        var sheet = document.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().Single(candidate => candidate.Name!.Value == sheetName);
        var worksheet = Worksheet(document, sheet);
        var cell = worksheet.GetFirstChild<SheetData>()!.Elements<Row>()
            .Skip(1)
            .Single()
            .Elements<Cell>()
            .First();
        cell.RemoveAllChildren();
        cell.DataType = CellValues.InlineString;
        cell.InlineString = new InlineString(new Text(value));
        worksheet.Save();
    }

    private static string TemporaryPath() => Path.Combine(Path.GetTempPath(), $"cellsharp-workbook-{Guid.NewGuid():N}.xlsx");

    private static void Delete(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }

    private sealed class Customer
    {
        public string? Name { get; set; }

        public Status? Status { get; set; }

        public string? Category { get; set; }
    }

    private sealed class Order
    {
        public int Id { get; set; }

        public decimal Amount { get; set; }

        public string? Category { get; set; }
    }

    private sealed class Status : IEquatable<Status>
    {
        internal static Status Active { get; } = new("ACTIVE");

        internal Status(string code) => Code = code;

        internal string Code { get; }

        public bool Equals(Status? other) => other is not null && Code == other.Code;

        public override bool Equals(object? obj) => Equals(obj as Status);

        public override int GetHashCode() => Code.GetHashCode();
    }

    private sealed class StatusConverter : IExcelValueConverter<Status?, string>
    {
        internal static StatusConverter Instance { get; } = new();

        public string Write(Status? value) => value?.Code ?? string.Empty;

        public bool TryRead(string value, out Status? converted)
        {
            converted = value == "ACTIVE" ? Status.Active : null;
            return converted is not null;
        }
    }
}
