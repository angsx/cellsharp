using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelAttributeSchemaTests
{
    [Fact]
    public void AttributeSchemaAppliesHeadersOrderAndConventionFallback()
    {
        var path = TemporaryPath();
        var schema = Excel.SchemaFromAttributes<AttributedCustomer>();

        try
        {
            Excel.Write(path, [new AttributedCustomer { Id = 7, Name = "Ada", Note = "Preferred", Hidden = "secret" }], schema);

            using var document = SpreadsheetDocument.Open(path, false);
            var headers = HeaderCells(document).Select(Text);

            Assert.Equal(["Customer ID", "Customer name", "Note"], headers);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void AttributeSchemaIgnoreOptionalAndFormatUseTheNormalSchemaPipeline()
    {
        var path = TemporaryPath();
        var templatePath = TemporaryPath();
        var schema = Excel.SchemaFromAttributes<FormattedCustomer>();

        try
        {
            Excel.Write(path, [new FormattedCustomer { Amount = 12.5M, OptionalName = "Ada", Hidden = "secret" }], schema);
            Excel.CreateTemplate(templatePath, schema);

            using var document = SpreadsheetDocument.Open(path, false);
            var value = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.GetFirstChild<SheetData>()!
                .Elements<Row>().Skip(1).Single().Elements<Cell>().First();
            var format = CellFormat(document, value);
            var numberFormat = document.WorkbookPart.WorkbookStylesPart!.Stylesheet!.NumberingFormats!
                .Elements<NumberingFormat>()
                .Single(numberFormat => numberFormat.NumberFormatId!.Value == format.NumberFormatId!.Value);

            Assert.Equal("#,##0.00", numberFormat.FormatCode!.Value);

            using var template = SpreadsheetDocument.Open(templatePath, false);
            Assert.Equal(["Amount", "OptionalName"], HeaderCells(template).Select(Text));

            CreateWorkbook(path, ["Amount"], ["5.0"]);
            var imported = Excel.Read<FormattedCustomer>(path, schema);
            Assert.True(imported.IsValid);
            Assert.Null(Assert.Single(imported.Rows).OptionalName);
        }
        finally
        {
            Delete(path);
            Delete(templatePath);
        }
    }

    [Fact]
    public void AttributeSchemaOrdersExplicitColumnsBeforeUnorderedAndKeepsOrderTiesStable()
    {
        var path = TemporaryPath();

        try
        {
            Excel.Write(path, [new OrderedCustomer()], Excel.SchemaFromAttributes<OrderedCustomer>());

            using var document = SpreadsheetDocument.Open(path, false);
            Assert.Equal(["Also first", "First", "Unordered", "AnotherUnordered"], HeaderCells(document).Select(Text));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void FluentConfigurationOverridesAttributeDefaultsWithoutDuplicatingTheColumn()
    {
        var path = TemporaryPath();
        var schema = Excel.SchemaFromAttributes<OverrideCustomer>(builder => builder
            .Column(customer => customer.Name, column => column
                .Header("Override")
                .MapFromHeader("Input name")
                .Validate(name => name.Length > 1, "Name is too short")));

        try
        {
            Excel.Write(path, [new OverrideCustomer { Name = "Ada" }], schema);
            using (var document = SpreadsheetDocument.Open(path, false))
            {
                Assert.Equal(["Override"], HeaderCells(document).Select(Text));
            }

            CreateWorkbook(path, ["Input name"], ["A"]);
            var imported = Excel.Read<OverrideCustomer>(path, schema);
            Assert.Empty(imported.Rows);
            Assert.Equal("Name is too short", Assert.Single(imported.Errors).Message);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void FluentConfigurationCanAddAConverterToAnAttributeColumn()
    {
        var path = TemporaryPath();
        var schema = Excel.SchemaFromAttributes<ConvertedCustomer>(builder => builder
            .Column(customer => customer.Status, column => column
                .ConvertWith(CustomerStatusConverter.Instance)
                .MapFromHeader("Imported status")
                .Validate(status => status == CustomerStatus.Active, "Only active customers are accepted")));

        try
        {
            Excel.Write(path, [new ConvertedCustomer { Status = CustomerStatus.Active }], schema);
            using (var document = SpreadsheetDocument.Open(path, false))
            {
                var value = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.GetFirstChild<SheetData>()!
                    .Elements<Row>().Skip(1).Single().Elements<Cell>().Single();
                Assert.Equal("ACTIVE", Text(value));
            }

            CreateWorkbook(path, ["Imported status"], ["INACTIVE"]);
            var imported = Excel.Read<ConvertedCustomer>(path, schema);
            Assert.Empty(imported.Rows);
            Assert.Equal("Only active customers are accepted", Assert.Single(imported.Errors).Message);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void SchemaLessApisDoNotInspectCellSharpAttributes()
    {
        var path = TemporaryPath();

        try
        {
            Excel.Write(path, [new AttributedCustomer { Id = 1, Name = "Ada", Note = "Note", Hidden = "Visible" }]);

            using (var document = SpreadsheetDocument.Open(path, false))
            {
                Assert.Equal(["Id", "Name", "Note", "Hidden"], HeaderCells(document).Select(Text));
            }

            CreateWorkbook(path, ["Id", "Name", "Note", "Hidden"], ["2", "Grace", "Imported", "Also imported"]);
            var imported = Excel.Read<AttributedCustomer>(path);
            Assert.Equal("Also imported", Assert.Single(imported.Rows).Hidden);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void InvalidAttributeCombinationsIdentifyThePropertyAndType()
    {
        var exception = Assert.Throws<ArgumentException>(() => Excel.SchemaFromAttributes<ConflictingAttributes>());

        Assert.Contains(nameof(ConflictingAttributes.Value), exception.Message);
        Assert.Contains(typeof(ConflictingAttributes).FullName!, exception.Message);
    }

    private static IEnumerable<Cell> HeaderCells(SpreadsheetDocument document) => document.WorkbookPart!.WorksheetParts.Single().Worksheet!
        .GetFirstChild<SheetData>()!.Elements<Row>().First().Elements<Cell>();

    private static CellFormat CellFormat(SpreadsheetDocument document, Cell cell) => document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!
        .CellFormats!.Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);

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
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1U, Name = "Sheet1" });
        worksheetPart.Worksheet.Save();
        workbookPart.Workbook.Save();
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

    private static string Text(Cell cell) => cell.InlineString?.Text?.Text ?? cell.CellValue?.Text ?? string.Empty;

    private static string TemporaryPath() => Path.Combine(Path.GetTempPath(), $"cellsharp-attributes-{Guid.NewGuid():N}.xlsx");

    private static void Delete(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }

    private sealed class AttributedCustomer
    {
        [ExcelColumn("Customer ID", Order = 1)]
        public int Id { get; set; }

        [ExcelColumn("Customer name", Order = 2)]
        public string? Name { get; set; }

        public string? Note { get; set; }

        [ExcelIgnore]
        public string? Hidden { get; set; }
    }

    private sealed class FormattedCustomer
    {
        [ExcelColumn(Format = "#,##0.00")]
        public decimal Amount { get; set; }

        [ExcelColumn(Optional = true)]
        public string? OptionalName { get; set; }

        [ExcelIgnore]
        public string? Hidden { get; set; }
    }

    private sealed class OrderedCustomer
    {
        public string? Unordered { get; set; }

        [ExcelColumn("Also first", Order = 1)]
        public string? AlsoFirst { get; set; }

        [ExcelColumn("First", Order = 1)]
        public string? First { get; set; }

        public string? AnotherUnordered { get; set; }
    }

    private sealed class OverrideCustomer
    {
        [ExcelColumn("Original")]
        public string Name { get; set; } = string.Empty;
    }

    private sealed class ConvertedCustomer
    {
        [ExcelColumn("Status")]
        public CustomerStatus Status { get; set; }
    }

    private enum CustomerStatus
    {
        Active,
        Inactive,
    }

    private sealed class CustomerStatusConverter : IExcelValueConverter<CustomerStatus, string>
    {
        internal static CustomerStatusConverter Instance { get; } = new();

        public string Write(CustomerStatus value) => value == CustomerStatus.Active ? "ACTIVE" : "INACTIVE";

        public bool TryRead(string value, out CustomerStatus converted)
        {
            if (value == "ACTIVE")
            {
                converted = CustomerStatus.Active;
                return true;
            }

            if (value == "INACTIVE")
            {
                converted = CustomerStatus.Inactive;
                return true;
            }

            converted = default;
            return false;
        }
    }

    private sealed class ConflictingAttributes
    {
        [ExcelColumn("Value")]
        [ExcelIgnore]
        public string? Value { get; set; }
    }
}
