using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelConverterTests
{
    private static readonly string[] CustomerHeaders = ["Name", "Status"];

    [Fact]
    public void ConverterWritesReadsAndRoundTripsAnUnsupportedDomainType()
    {
        var path = TemporaryPath();
        var schema = CustomerSchema(StatusConverter.Instance);

        try
        {
            Excel.Write(path, [new ConverterCustomer { Name = "Ada", Status = CustomerStatus.Active }], schema);

            using var document = SpreadsheetDocument.Open(path, false);
            var cells = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.GetFirstChild<SheetData>()!
                .Elements<Row>()
                .Skip(1)
                .Single()
                .Elements<Cell>()
                .ToArray();
            var result = Excel.Read<ConverterCustomer>(path, schema);

            Assert.Equal("ACTIVE", Text(cells[1]));
            Assert.Equal(CustomerStatus.Active, Assert.Single(result.Rows).Status);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ConverterFailureUsesInvalidValueAndPreservesOtherRows()
    {
        var path = TemporaryPath();
        var schema = CustomerSchema(StatusConverter.Instance);

        try
        {
            CreateWorkbook(path, CustomerHeaders, ["Ada", "ACTIVE"], ["Grace", "ARCH"]);

            var result = Excel.Read<ConverterCustomer>(path, schema);

            Assert.Equal(CustomerStatus.Active, Assert.Single(result.Rows).Status);
            var error = Assert.Single(result.Errors);
            Assert.Equal(ExcelReadErrorCode.InvalidValue, error.Code);
            Assert.Equal("Value 'ARCH' could not be converted to 'CellSharp.Tests.ExcelConverterTests+CustomerStatus'.", error.Message);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void UnsupportedTypesRequireAConverterButTheConverterCanBeReused()
    {
        Assert.Throws<NotSupportedException>(() => Excel.Schema<ConverterCustomer>()
            .Column(customer => customer.Status));

        var firstPath = TemporaryPath();
        var secondPath = TemporaryPath();
        var schema = CustomerSchema(StatusConverter.Instance);
        try
        {
            Excel.Write(firstPath, [new ConverterCustomer { Name = "Ada", Status = CustomerStatus.Active }], schema);
            Excel.Write(secondPath, [new ConverterCustomer { Name = "Grace", Status = CustomerStatus.Inactive }], schema);

            Assert.Equal(CustomerStatus.Active, Assert.Single(Excel.Read<ConverterCustomer>(firstPath, schema).Rows).Status);
            Assert.Equal(CustomerStatus.Inactive, Assert.Single(Excel.Read<ConverterCustomer>(secondPath, schema).Rows).Status);
        }
        finally
        {
            Delete(firstPath);
            Delete(secondPath);
        }
    }

    [Fact]
    public void ConverterRunsBeforeCustomValidation()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<ConverterCustomer>()
            .Column(customer => customer.Name)
            .Column(customer => customer.Status, column => column
                .ConvertWith(StatusConverter.Instance)
                .Validate(status => status != CustomerStatus.Archived, "Archived status is not allowed"))
            .Build();

        try
        {
            Excel.Write(path, [new ConverterCustomer { Name = "Ada", Status = CustomerStatus.Archived }], schema);

            var result = Excel.Read<ConverterCustomer>(path, schema);

            Assert.Empty(result.Rows);
            Assert.Equal("Archived status is not allowed", Assert.Single(result.Errors).Message);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ConverterOutputCanUseNumericFormatsAndTemplateColumnStyles()
    {
        var path = TemporaryPath();
        var templatePath = TemporaryPath();
        var schema = Excel.Schema<Invoice>()
            .Column(invoice => invoice.Total, column => column
                .ConvertWith(MoneyConverter.Instance)
                .Format("#,##0.00"))
            .Build();

        try
        {
            Excel.Write(path, [new Invoice { Total = new Money(1234.5M) }], schema);
            Excel.CreateTemplate(templatePath, schema);

            using var document = SpreadsheetDocument.Open(path, false);
            var value = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.GetFirstChild<SheetData>()!
                .Elements<Row>()
                .Skip(1)
                .Single()
                .Elements<Cell>()
                .Single();
            using var template = SpreadsheetDocument.Open(templatePath, false);
            var column = template.WorkbookPart!.WorksheetParts.Single().Worksheet!.GetFirstChild<Columns>()!
                .Elements<Column>()
                .Single();

            Assert.Equal("1234.5", value.CellValue!.Text);
            Assert.Equal("#,##0.00", NumberFormat(document, CellFormat(document, value)));
            Assert.Equal("#,##0.00", NumberFormat(template, ColumnFormat(template, column)));
        }
        finally
        {
            Delete(path);
            Delete(templatePath);
        }
    }

    [Fact]
    public void NullableConverterColumnsBypassEmptyCellsAndOptionalHeaders()
    {
        var blankPath = TemporaryPath();
        var missingPath = TemporaryPath();
        var schema = Excel.Schema<NullableConverterCustomer>()
            .Column(customer => customer.Name)
            .Column(customer => customer.Status, column => column
                .Optional()
                .ConvertWith(StatusConverter.Instance))
            .Build();

        try
        {
            Excel.Write(blankPath, [new NullableConverterCustomer { Name = "Ada", Status = null }], schema);
            CreateWorkbook(missingPath, ["Name"], ["Grace"]);

            Assert.Null(Assert.Single(Excel.Read<NullableConverterCustomer>(blankPath, schema).Rows).Status);
            Assert.Null(Assert.Single(Excel.Read<NullableConverterCustomer>(missingPath, schema).Rows).Status);
        }
        finally
        {
            Delete(blankPath);
            Delete(missingPath);
        }
    }

    [Fact]
    public void ConverterCanCoexistWithCompatibleNativeDataValidationAndStyling()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<CodeCustomer>()
            .Column(customer => customer.Code, column => column
                .ConvertWith(UppercaseStringConverter.Instance)
                .AllowedValues("ACTIVE", "INACTIVE"))
            .Build();

        try
        {
            Excel.Write(path, [new CodeCustomer { Code = "active" }], schema, options => options.Theme(ExcelTheme.Modern));

            using var document = SpreadsheetDocument.Open(path, false);
            var worksheet = document.WorkbookPart!.WorksheetParts.First().Worksheet!;
            var value = worksheet.GetFirstChild<SheetData>()!.Elements<Row>().Skip(1).Single().Elements<Cell>().Single();
            var validation = worksheet.GetFirstChild<DataValidations>()!.Elements<DataValidation>().Single();

            Assert.Equal("ACTIVE", Text(value));
            Assert.Equal(DataValidationValues.List, validation.Type!.Value);
            Assert.NotEqual(0U, value.StyleIndex!.Value);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ConverterExceptionsPropagateInsteadOfBecomingWorkbookDataErrors()
    {
        var path = TemporaryPath();
        var schema = CustomerSchema(ThrowingStatusConverter.Instance);

        try
        {
            CreateWorkbook(path, CustomerHeaders, ["Ada", "ACTIVE"]);

            Assert.Throws<InvalidOperationException>(() => Excel.Read<ConverterCustomer>(path, schema));
        }
        finally
        {
            Delete(path);
        }
    }

    private static ExcelSchema<ConverterCustomer> CustomerSchema(IExcelValueConverter<CustomerStatus?, string> converter) => Excel.Schema<ConverterCustomer>()
        .Column(customer => customer.Name)
        .Column(customer => customer.Status, column => column.ConvertWith(converter))
        .Build();

    private static void CreateWorkbook(string path, string[] headers, params string[][] values)
    {
        using var document = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var data = new SheetData();
        data.AppendChild(Row(headers));
        foreach (var value in values)
        {
            data.AppendChild(Row(value));
        }

        worksheetPart.Worksheet = new Worksheet(data);
        var sheets = workbookPart.Workbook!.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1U, Name = "Sheet1" });
        worksheetPart.Worksheet!.Save();
        workbookPart.Workbook!.Save();
    }

    private static Row Row(IEnumerable<string> values)
    {
        var row = new Row();
        foreach (var value in values)
        {
            row.AppendChild(new Cell
            {
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new Text(value)),
            });
        }

        return row;
    }

    private static CellFormat CellFormat(SpreadsheetDocument document, Cell cell) => document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!
        .CellFormats!
        .Elements<CellFormat>()
        .ElementAt((int)cell.StyleIndex!.Value);

    private static CellFormat ColumnFormat(SpreadsheetDocument document, Column column) => document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!
        .CellFormats!
        .Elements<CellFormat>()
        .ElementAt((int)column.Style!.Value);

    private static string? NumberFormat(SpreadsheetDocument document, CellFormat format) => document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!
        .NumberingFormats!
        .Elements<NumberingFormat>()
        .Single(numberFormat => numberFormat.NumberFormatId!.Value == format.NumberFormatId!.Value)
        .FormatCode!
        .Value;

    private static string Text(Cell cell) => cell.InlineString?.Text?.Text ?? cell.CellValue?.Text ?? string.Empty;

    private static string TemporaryPath() => Path.Combine(Path.GetTempPath(), $"cellsharp-converter-{Guid.NewGuid():N}.xlsx");

    private static void Delete(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }

    private sealed class ConverterCustomer
    {
        public string? Name { get; set; }

        public CustomerStatus? Status { get; set; } = CustomerStatus.Active;
    }

    private sealed class NullableConverterCustomer
    {
        public string? Name { get; set; }

        public CustomerStatus? Status { get; set; }
    }

    private sealed class CodeCustomer
    {
        public string? Code { get; set; }
    }

    private sealed class Invoice
    {
        public Money Total { get; set; } = new(0M);
    }

    private sealed class CustomerStatus : IEquatable<CustomerStatus>
    {
        internal static CustomerStatus Active { get; } = new("ACTIVE");
        internal static CustomerStatus Inactive { get; } = new("INACTIVE");
        internal static CustomerStatus Archived { get; } = new("ARCHIVED");

        internal CustomerStatus(string code)
        {
            Code = code;
        }

        internal string Code { get; }

        public bool Equals(CustomerStatus? other) => other is not null && Code == other.Code;

        public override bool Equals(object? obj) => Equals(obj as CustomerStatus);

        public override int GetHashCode() => Code.GetHashCode();
    }

    private sealed class Money
    {
        internal Money(decimal amount)
        {
            Amount = amount;
        }

        internal decimal Amount { get; }
    }

    private sealed class StatusConverter : IExcelValueConverter<CustomerStatus?, string>
    {
        internal static StatusConverter Instance { get; } = new();

        public string Write(CustomerStatus? value) => value?.Code ?? string.Empty;

        public bool TryRead(string value, out CustomerStatus? converted)
        {
            converted = value switch
            {
                "ACTIVE" => CustomerStatus.Active,
                "INACTIVE" => CustomerStatus.Inactive,
                "ARCHIVED" => CustomerStatus.Archived,
                _ => CustomerStatus.Active,
            };
            return value == "ACTIVE" || value == "INACTIVE" || value == "ARCHIVED";
        }
    }

    private sealed class MoneyConverter : IExcelValueConverter<Money, decimal>
    {
        internal static MoneyConverter Instance { get; } = new();

        public decimal Write(Money value) => value.Amount;

        public bool TryRead(decimal value, out Money converted)
        {
            converted = new Money(value);
            return true;
        }
    }

    private sealed class ThrowingStatusConverter : IExcelValueConverter<CustomerStatus?, string>
    {
        internal static ThrowingStatusConverter Instance { get; } = new();

        public string Write(CustomerStatus? value) => value?.Code ?? string.Empty;

        public bool TryRead(string value, out CustomerStatus? converted)
        {
            throw new InvalidOperationException("Converter failure.");
        }
    }

    private sealed class UppercaseStringConverter : IExcelValueConverter<string?, string>
    {
        internal static UppercaseStringConverter Instance { get; } = new();

        public string Write(string? value) => value?.ToUpperInvariant() ?? string.Empty;

        public bool TryRead(string value, out string? converted)
        {
            converted = value.ToUpperInvariant();
            return true;
        }
    }
}
