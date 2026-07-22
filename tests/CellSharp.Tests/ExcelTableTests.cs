using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelTableTests
{
    [Fact]
    public void WriteCreatesANativeTableWithExactRangeColumnsFiltersAndStyle()
    {
        var path = TemporaryPath();
        var schema = CustomerSchema("Customers").AsTable("CustomersTable", "TableStyleLight9").Build();

        try
        {
            Excel.Write(path,
            [
                new Customer { Id = 1, Name = "Ada", Status = "Active" },
                new Customer { Id = 2, Name = "Grace", Status = "Inactive" },
            ], schema);

            using var document = SpreadsheetDocument.Open(path, false);
            var worksheetPart = document.WorkbookPart!.WorksheetParts.Single();
            var part = Assert.Single(worksheetPart.TableDefinitionParts);
            var table = part.Table!;
            var tablePart = Assert.Single(worksheetPart.Worksheet!.GetFirstChild<TableParts>()!.Elements<TablePart>());

            Assert.Equal("A1:C3", table.Reference!.Value);
            Assert.Equal(1U, table.Id!.Value);
            Assert.Equal("CustomersTable", table.Name!.Value);
            Assert.Equal("CustomersTable", table.DisplayName!.Value);
            Assert.Equal("A1:C3", table.AutoFilter!.Reference!.Value);
            Assert.Equal(3U, table.TableColumns!.Count!.Value);
            Assert.Equal(["Id", "Name", "Status"], table.TableColumns.Elements<TableColumn>().Select(column => column.Name!.Value));
            Assert.Equal(worksheetPart.GetIdOfPart(part), tablePart.Id!.Value);
            Assert.Equal("TableStyleLight9", table.TableStyleInfo!.Name!.Value);
            Assert.True(table.TableStyleInfo.ShowRowStripes!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void AutomaticTableNamesAreDeterministicAndGloballyUnique()
    {
        var path = TemporaryPath();
        var first = CustomerSchema("Orders Archive").AsTable().Build();
        var second = CustomerSchema("Orders-Archive").AsTable().Build();

        try
        {
            Excel.Workbook()
                .AddSheet([new Customer { Id = 1, Name = "Ada", Status = "Active" }], first)
                .AddSheet([new Customer { Id = 2, Name = "Grace", Status = "Inactive" }], second)
                .Write(path);

            using var document = SpreadsheetDocument.Open(path, false);
            var tables = document.WorkbookPart!.WorksheetParts
                .SelectMany(part => part.TableDefinitionParts)
                .Select(part => part.Table!)
                .OrderBy(table => table.Id!.Value)
                .ToArray();

            Assert.Equal(["OrdersArchiveTable", "OrdersArchiveTable2"], tables.Select(table => table.Name!.Value));
            Assert.Equal([1U, 2U], tables.Select(table => table.Id!.Value));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void TableUsesEffectiveHeadersAndExcludesIgnoredColumns()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<Customer>()
            .AsTable("PeopleTable")
            .Column(customer => customer.Name, column => column.Header("Customer name"))
            .Column(customer => customer.Id)
            .Column(customer => customer.Status, column => column.Ignore())
            .Build();

        try
        {
            Excel.Write(path, [new Customer { Id = 1, Name = "Ada", Status = "Active" }], schema);

            using var document = SpreadsheetDocument.Open(path, false);
            var table = document.WorkbookPart!.WorksheetParts.Single().TableDefinitionParts.Single().Table!;

            Assert.Equal("A1:B2", table.Reference!.Value);
            Assert.Equal(["Customer name", "Id"], table.TableColumns!.Elements<TableColumn>().Select(column => column.Name!.Value));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void EmptyTableExportsTheHeaderOnlyRangeAndTemplatesUseTheSameSemantics()
    {
        var dataPath = TemporaryPath();
        var templatePath = TemporaryPath();
        var schema = CustomerSchema("Customers").AsTable().Build();

        try
        {
            Excel.Write(dataPath, Array.Empty<Customer>(), schema);
            Excel.CreateTemplate(templatePath, schema);

            using var data = SpreadsheetDocument.Open(dataPath, false);
            using var template = SpreadsheetDocument.Open(templatePath, false);

            Assert.Equal("A1:C1", data.WorkbookPart!.WorksheetParts.Single().TableDefinitionParts.Single().Table!.Reference!.Value);
            Assert.Equal("A1:C1", template.WorkbookPart!.WorksheetParts.Single().TableDefinitionParts.Single().Table!.Reference!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(data));
            Assert.Empty(new OpenXmlValidator().Validate(template));
        }
        finally
        {
            Delete(dataPath);
            Delete(templatePath);
        }
    }

    [Fact]
    public void TablePreservesFormulaCellsValidationDirectFormattingAndLiteralFormulaLookingText()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<Invoice>()
            .AsTable("InvoicesTable")
            .Column(invoice => invoice.Name)
            .Column(invoice => invoice.Quantity)
            .Column(invoice => invoice.UnitPrice, column => column.AllowedValues("10", "12.5"))
            .Column(invoice => invoice.Total, column => column
                .Formula(context => $"=B{context.Row}*C{context.Row}")
                .Format("#,##0.00")
                .Width(18D)
                .Align(ExcelHorizontalAlignment.Right))
            .Build();

        try
        {
            Excel.Write(path, [new Invoice { Name = "=SUM(A1:A2)", Quantity = 2, UnitPrice = "12.5", Total = 0M }], schema);

            using var document = SpreadsheetDocument.Open(path, false);
            var worksheet = document.WorkbookPart!.WorksheetParts.First().Worksheet!;
            var cells = worksheet.GetFirstChild<SheetData>()!.Elements<Row>().Skip(1).Single().Elements<Cell>().ToArray();
            var table = document.WorkbookPart.WorksheetParts.First().TableDefinitionParts.Single().Table!;
            var totalColumn = worksheet.GetFirstChild<Columns>()!.Elements<Column>().Single(column => column.Min!.Value == 4U);
            var totalFormat = document.WorkbookPart.WorkbookStylesPart!.Stylesheet!.CellFormats!
                .Elements<CellFormat>()
                .ElementAt((int)cells[3].StyleIndex!.Value);
            var totalNumberFormat = document.WorkbookPart.WorkbookStylesPart.Stylesheet!.NumberingFormats!
                .Elements<NumberingFormat>()
                .Single(format => format.NumberFormatId!.Value == totalFormat.NumberFormatId!.Value);

            Assert.Equal("A1:D2", table.Reference!.Value);
            Assert.Equal(CellValues.InlineString, cells[0].DataType!.Value);
            Assert.Equal("=SUM(A1:A2)", cells[0].InlineString!.Text!.Text);
            Assert.Null(cells[0].CellFormula);
            Assert.Equal("B2*C2", cells[3].CellFormula!.InnerText);
            Assert.NotNull(worksheet.GetFirstChild<DataValidations>());
            Assert.NotNull(document.WorkbookPart.Workbook!.CalculationProperties);
            Assert.NotEqual(0U, cells[3].StyleIndex!.Value);
            Assert.Equal(18D, totalColumn.Width!.Value);
            Assert.Equal(HorizontalAlignmentValues.Right, totalFormat.Alignment!.Horizontal!.Value);
            Assert.Equal("#,##0.00", totalNumberFormat.FormatCode!.Value);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void TableRoundTripsThroughTheExistingWorksheetReader()
    {
        var path = TemporaryPath();
        var schema = CustomerSchema("Customers").AsTable().Build();

        try
        {
            Excel.Write(path, [new Customer { Id = 7, Name = "Ada", Status = "Active" }], schema);

            var result = Excel.Read<Customer>(path, schema);

            var customer = Assert.Single(result.Rows);
            Assert.True(result.IsValid);
            Assert.Equal(7, customer.Id);
            Assert.Equal("Ada", customer.Name);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void TableNamesFailFastForInvalidAndDuplicateExplicitNames()
    {
        Assert.Throws<ArgumentException>(() => CustomerSchema("Customers").AsTable("A1"));
        Assert.Throws<ArgumentException>(() => CustomerSchema("Customers").AsTable("Customer Table"));
        Assert.Throws<ArgumentException>(() => Excel.Schema<Customer>()
            .AsTable()
            .Column(customer => customer.Id, column => column.Header("Value"))
            .Column(customer => customer.Name, column => column.Header("value")));

        var path = TemporaryPath();
        var first = CustomerSchema("First").AsTable("SharedTable").Build();
        var second = CustomerSchema("Second").AsTable("sharedtable").Build();
        try
        {
            Assert.Throws<InvalidOperationException>(() => Excel.Workbook()
                .AddSheet(Array.Empty<Customer>(), first)
                .AddSheet(Array.Empty<Customer>(), second)
                .Write(path));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void TableWriteKeepsStreamOwnershipAndCooperativeCancellation()
    {
        var schema = CustomerSchema("Customers").AsTable().Build();
        using var stream = new TrackingStream();

        Excel.Write(stream, [new Customer { Id = 1, Name = "Ada", Status = "Active" }], schema);

        using var source = new CancellationTokenSource();
        Assert.Throws<OperationCanceledException>(() => Excel.Write(stream, RowsThatCancel(source), schema, source.Token));
        Assert.False(stream.WasDisposed);
    }

    private static ExcelSchemaBuilder<Customer> CustomerSchema(string sheetName) => Excel.Schema<Customer>()
        .SheetName(sheetName)
        .Column(customer => customer.Id)
        .Column(customer => customer.Name)
        .Column(customer => customer.Status);

    private static IEnumerable<Customer> RowsThatCancel(CancellationTokenSource source)
    {
        yield return new Customer { Id = 1, Name = "Ada", Status = "Active" };
        source.Cancel();
        yield return new Customer { Id = 2, Name = "Grace", Status = "Inactive" };
    }

    private static string TemporaryPath() => Path.Combine(Path.GetTempPath(), $"cellsharp-table-{Guid.NewGuid():N}.xlsx");

    private static void Delete(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }

    private sealed class Customer
    {
        public int Id { get; set; }

        public string? Name { get; set; }

        public string? Status { get; set; }
    }

    private sealed class Invoice
    {
        public string? Name { get; set; }

        public int Quantity { get; set; }

        public string? UnitPrice { get; set; }

        public decimal Total { get; set; }
    }

    private sealed class TrackingStream : MemoryStream
    {
        internal bool WasDisposed { get; private set; }

        protected override void Dispose(bool disposing)
        {
            WasDisposed = true;
            base.Dispose(disposing);
        }
    }
}
