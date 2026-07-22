using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelWorksheetSettingsTests
{
    [Fact]
    public void StandaloneAutoFilterUsesTheEffectiveDataRangeForDataEmptyTemplatesAndOverlays()
    {
        var dataPath = TemporaryPath();
        var emptyPath = TemporaryPath();
        var templatePath = TemporaryPath();
        var overlayPath = TemporaryPath();
        var schema = Schema("Products").AutoFilter().Build();
        var overlay = Excel.Overlay<Product>(runtime => runtime
            .Include(product => product.Name, false)
            .Include(product => product.Category, false));

        try
        {
            Excel.Write(dataPath, [new Product { Id = 1, Name = "A", Category = "Retail" }, new Product { Id = 2, Name = "B", Category = "Wholesale" }], schema);
            Excel.Write(emptyPath, Array.Empty<Product>(), schema);
            Excel.CreateTemplate(templatePath, schema);
            Excel.Write(overlayPath, [new Product { Id = 1, Name = "A", Category = "Retail" }], schema, overlay);

            using var data = SpreadsheetDocument.Open(dataPath, false);
            using var empty = SpreadsheetDocument.Open(emptyPath, false);
            using var template = SpreadsheetDocument.Open(templatePath, false);
            using var overlayDocument = SpreadsheetDocument.Open(overlayPath, false);

            Assert.Equal("A1:C3", Worksheet(data).GetFirstChild<AutoFilter>()!.Reference!.Value);
            Assert.Equal("A1:C1", Worksheet(empty).GetFirstChild<AutoFilter>()!.Reference!.Value);
            Assert.Equal("A1:C1", Worksheet(template).GetFirstChild<AutoFilter>()!.Reference!.Value);
            Assert.Equal("A1:A2", Worksheet(overlayDocument).GetFirstChild<AutoFilter>()!.Reference!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(data));
            Assert.Empty(new OpenXmlValidator().Validate(template));
        }
        finally
        {
            Delete(dataPath);
            Delete(emptyPath);
            Delete(templatePath);
            Delete(overlayPath);
        }
    }

    [Fact]
    public void TableOwnsTheFilterWhenAutoFilterIsAlsoConfigured()
    {
        var path = TemporaryPath();
        var schema = Schema("Products").AsTable().AutoFilter().Build();

        try
        {
            Excel.Write(path, [new Product { Id = 1, Name = "A", Category = "Retail" }], schema);

            using var document = SpreadsheetDocument.Open(path, false);
            var part = document.WorkbookPart!.WorksheetParts.Single();

            Assert.Null(part.Worksheet!.GetFirstChild<AutoFilter>());
            Assert.Equal("A1:C2", part.TableDefinitionParts.Single().Table!.AutoFilter!.Reference!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void FreezePanesWritesCorrectPanesAndPreservesFreezeHeaderRowCompatibility()
    {
        var rowPath = TemporaryPath();
        var columnPath = TemporaryPath();
        var combinedPath = TemporaryPath();
        var legacyPath = TemporaryPath();
        var nonePath = TemporaryPath();

        try
        {
            Excel.Write(rowPath, Array.Empty<Product>(), Schema("Rows").FreezePanes(1, 0).Build());
            Excel.Write(columnPath, Array.Empty<Product>(), Schema("Columns").FreezePanes(0, 1).Build());
            Excel.Write(combinedPath, Array.Empty<Product>(), Schema("Both").FreezePanes(2, 3).Build());
            Excel.Write(legacyPath, Array.Empty<Product>(), options => options.FreezeHeaderRow());
            Excel.Write(nonePath, Array.Empty<Product>(), Schema("None").FreezePanes(0, 0).Build());

            using var rows = SpreadsheetDocument.Open(rowPath, false);
            using var columns = SpreadsheetDocument.Open(columnPath, false);
            using var both = SpreadsheetDocument.Open(combinedPath, false);
            using var legacy = SpreadsheetDocument.Open(legacyPath, false);
            using var none = SpreadsheetDocument.Open(nonePath, false);

            AssertPane(Worksheet(rows), 1D, null, "A2", PaneValues.BottomLeft);
            AssertPane(Worksheet(columns), null, 1D, "B1", PaneValues.TopRight);
            AssertPane(Worksheet(both), 2D, 3D, "D3", PaneValues.BottomRight);
            AssertPane(Worksheet(legacy), 1D, null, "A2", PaneValues.BottomLeft);
            Assert.Null(Worksheet(none).GetFirstChild<SheetViews>());
        }
        finally
        {
            Delete(rowPath);
            Delete(columnPath);
            Delete(combinedPath);
            Delete(legacyPath);
            Delete(nonePath);
        }
    }

    [Fact]
    public void WorksheetSettingsRejectNegativeFreezeAndFitValues()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => Schema("Products").FreezePanes(-1, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => Schema("Products").FreezePanes(0, -1));
        Assert.Throws<ArgumentOutOfRangeException>(() => Schema("Products").FitToPage(-1, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => Schema("Products").FitToPage(0, -1));
    }

    [Fact]
    public void PrintSettingsWriteNativePageSetupPrintOptionsAndEscapedPrintTitles()
    {
        var path = TemporaryPath();
        var schema = Schema("O'Brien")
            .Landscape()
            .FitToPage(1, 0)
            .RepeatHeaderRowOnPrint()
            .PrintGridlines()
            .Build();

        try
        {
            Excel.Write(path, [new Product { Id = 1, Name = "A", Category = "Retail" }], schema);

            using var document = SpreadsheetDocument.Open(path, false);
            var worksheet = Worksheet(document);
            var pageSetup = worksheet.GetFirstChild<PageSetup>()!;
            var pageProperties = worksheet.GetFirstChild<SheetProperties>()!.PageSetupProperties!;
            var title = Assert.Single(document.WorkbookPart!.Workbook!.DefinedNames!.Elements<DefinedName>());

            Assert.Equal(OrientationValues.Landscape, pageSetup.Orientation!.Value);
            Assert.Equal(1U, pageSetup.FitToWidth!.Value);
            Assert.Equal(0U, pageSetup.FitToHeight!.Value);
            Assert.True(pageProperties.FitToPage!.Value);
            Assert.True(worksheet.GetFirstChild<PrintOptions>()!.GridLines!.Value);
            Assert.Equal("_xlnm.Print_Titles", title.Name!.Value);
            Assert.Equal(0U, title.LocalSheetId!.Value);
            Assert.Equal("'O''Brien'!$1:$1", title.Text);
            Assert.Null(worksheet.GetFirstChild<SheetViews>()?.Elements<SheetView>().FirstOrDefault()?.ShowGridLines);
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void FitToPageDoesNotSelectAnOrientationWhenTheSchemaDoesNotSelectOne()
    {
        var path = TemporaryPath();
        var schema = Schema("Products").FitToPage(1, 0).Build();

        try
        {
            Excel.Write(path, Array.Empty<Product>(), schema);

            using var document = SpreadsheetDocument.Open(path, false);
            var pageSetup = Worksheet(document).GetFirstChild<PageSetup>()!;

            Assert.Null(pageSetup.Orientation);
            Assert.Equal(1U, pageSetup.FitToWidth!.Value);
            Assert.Equal(0U, pageSetup.FitToHeight!.Value);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void MultiSheetSettingsAndPrintTitlesRemainIsolated()
    {
        var path = TemporaryPath();
        var customers = Schema("Customers").Portrait().Build();
        var orders = Schema("Order Details")
            .AutoFilter()
            .FreezePanes(1, 2)
            .Landscape()
            .FitToPage(1, 1)
            .RepeatHeaderRowOnPrint()
            .PrintGridlines()
            .Build();

        try
        {
            Excel.Workbook()
                .AddSheet([new Product { Id = 1, Name = "Ada", Category = "Retail" }], customers)
                .AddSheet([new Product { Id = 2, Name = "Order", Category = "Wholesale" }], orders)
                .Write(path);

            using var document = SpreadsheetDocument.Open(path, false);
            var customerWorksheet = Worksheet(document, "Customers");
            var orderWorksheet = Worksheet(document, "Order Details");
            var title = Assert.Single(document.WorkbookPart!.Workbook!.DefinedNames!.Elements<DefinedName>());

            Assert.Equal(OrientationValues.Portrait, customerWorksheet.GetFirstChild<PageSetup>()!.Orientation!.Value);
            Assert.Null(customerWorksheet.GetFirstChild<AutoFilter>());
            Assert.Null(customerWorksheet.GetFirstChild<PrintOptions>());
            Assert.Equal("A1:C2", orderWorksheet.GetFirstChild<AutoFilter>()!.Reference!.Value);
            AssertPane(orderWorksheet, 1D, 2D, "C2", PaneValues.BottomRight);
            Assert.Equal(OrientationValues.Landscape, orderWorksheet.GetFirstChild<PageSetup>()!.Orientation!.Value);
            Assert.Equal(1U, title.LocalSheetId!.Value);
            Assert.Equal("'Order Details'!$1:$1", title.Text);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void TemplatesKeepAllWorksheetSettingsWithoutAddingDataRows()
    {
        var path = TemporaryPath();
        var schema = Schema("Products")
            .AutoFilter()
            .FreezePanes(1, 1)
            .Landscape()
            .FitToPage(1, 1)
            .RepeatHeaderRowOnPrint()
            .PrintGridlines()
            .Build();

        try
        {
            Excel.CreateTemplate(path, schema);

            using var document = SpreadsheetDocument.Open(path, false);
            var worksheet = Worksheet(document);

            Assert.Single(worksheet.GetFirstChild<SheetData>()!.Elements<Row>());
            Assert.Equal("A1:C1", worksheet.GetFirstChild<AutoFilter>()!.Reference!.Value);
            AssertPane(worksheet, 1D, 1D, "B2", PaneValues.BottomRight);
            Assert.Equal(OrientationValues.Landscape, worksheet.GetFirstChild<PageSetup>()!.Orientation!.Value);
            Assert.True(worksheet.GetFirstChild<PrintOptions>()!.GridLines!.Value);
            Assert.Single(document.WorkbookPart!.Workbook!.DefinedNames!.Elements<DefinedName>());
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void SettingsIntegrateWithTablesFormulasValidationStreamsAndCancellation()
    {
        var schema = Excel.Schema<Invoice>()
            .SheetName("Invoices")
            .AsTable()
            .AutoFilter()
            .FreezePanes(1, 1)
            .Landscape()
            .FitToPage(1, 0)
            .RepeatHeaderRowOnPrint()
            .PrintGridlines()
            .Column(invoice => invoice.Name)
            .Column(invoice => invoice.Quantity)
            .Column(invoice => invoice.Category, column => column.AllowedValues("Retail", "Wholesale"))
            .Column(invoice => invoice.Total, column => column.Formula(context => $"=B{context.Row}*2"))
            .Build();
        using var stream = new TrackingStream();

        Excel.Write(stream, [new Invoice { Name = "=literal", Quantity = 2, Category = "Retail" }], schema);

        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = document.WorkbookPart!.WorksheetParts.First();
        var data = worksheetPart.Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().Skip(1).Single().Elements<Cell>().ToArray();
        Assert.Null(worksheetPart.Worksheet!.GetFirstChild<AutoFilter>());
        Assert.NotNull(worksheetPart.TableDefinitionParts.Single().Table!.AutoFilter);
        Assert.NotNull(worksheetPart.Worksheet!.GetFirstChild<DataValidations>());
        Assert.Equal("B2*2", data[3].CellFormula!.InnerText);
        Assert.Equal(CellValues.InlineString, data[0].DataType!.Value);
        Assert.NotNull(document.WorkbookPart.Workbook!.CalculationProperties);
        document.Dispose();

        using var source = new CancellationTokenSource();
        Assert.Throws<OperationCanceledException>(() => Excel.Write(stream, RowsThatCancel(source), schema, source.Token));
        Assert.False(stream.WasDisposed);
    }

    private static ExcelSchemaBuilder<Product> Schema(string sheetName) => Excel.Schema<Product>()
        .SheetName(sheetName)
        .Column(product => product.Id)
        .Column(product => product.Name)
        .Column(product => product.Category);

    private static Worksheet Worksheet(SpreadsheetDocument document) => document.WorkbookPart!.WorksheetParts.First().Worksheet!;

    private static Worksheet Worksheet(SpreadsheetDocument document, string name)
    {
        var sheet = document.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().Single(sheet => sheet.Name!.Value == name);
        return ((WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id!)).Worksheet!;
    }

    private static void AssertPane(Worksheet worksheet, double? vertical, double? horizontal, string topLeftCell, PaneValues activePane)
    {
        var pane = worksheet.GetFirstChild<SheetViews>()!.Elements<SheetView>().Single().GetFirstChild<Pane>()!;

        Assert.Equal(vertical, pane.VerticalSplit?.Value);
        Assert.Equal(horizontal, pane.HorizontalSplit?.Value);
        Assert.Equal(topLeftCell, pane.TopLeftCell!.Value);
        Assert.Equal(activePane, pane.ActivePane!.Value);
        Assert.Equal(PaneStateValues.Frozen, pane.State!.Value);
    }

    private static IEnumerable<Invoice> RowsThatCancel(CancellationTokenSource source)
    {
        yield return new Invoice { Name = "First", Quantity = 1, Category = "Retail" };
        source.Cancel();
        yield return new Invoice { Name = "Second", Quantity = 2, Category = "Wholesale" };
    }

    private static string TemporaryPath() => Path.Combine(Path.GetTempPath(), $"cellsharp-worksheet-settings-{Guid.NewGuid():N}.xlsx");

    private static void Delete(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }

    private sealed class Product
    {
        public int Id { get; set; }

        public string? Name { get; set; }

        public string? Category { get; set; }
    }

    private sealed class Invoice
    {
        public string? Name { get; set; }

        public int Quantity { get; set; }

        public string? Category { get; set; }

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
