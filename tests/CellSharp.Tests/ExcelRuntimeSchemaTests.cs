using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelRuntimeSchemaTests
{
    [Fact]
    public void OverlayChangesHeadersAndCompactsEnabledColumnsForExportAndTemplate()
    {
        var dataPath = TemporaryPath();
        var templatePath = TemporaryPath();
        var schema = ProductSchema();
        var overlay = Excel.Overlay<Product>(runtime => runtime
            .Header(product => product.Size, "Taglia")
            .Include(product => product.SerialNumber, false));

        try
        {
            Excel.Write(dataPath, [new Product { Name = "Shirt", Size = 42, SerialNumber = "SN-1" }], schema, overlay);
            Excel.CreateTemplate(templatePath, schema, overlay);

            Assert.Equal(["Name", "Taglia"], Headers(dataPath));
            Assert.Equal(["Name", "Taglia"], Headers(templatePath));
            using var template = SpreadsheetDocument.Open(templatePath, false);
            Assert.Empty(new OpenXmlValidator().Validate(template));
        }
        finally
        {
            Delete(dataPath);
            Delete(templatePath);
        }
    }

    [Fact]
    public void OverlayHeaderIsUsedForImportBeforeMapFromHeaderAndMatchesCaseInsensitively()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<Product>()
            .Column(product => product.Size, column => column.MapFromHeader("Legacy size"))
            .Build();
        var overlay = Excel.Overlay<Product>(runtime => runtime.Header(product => product.Size, "Taglia"));

        try
        {
            CreateWorkbook(path, ["tAgLiA"], ["42"]);
            var result = Excel.Read(path, schema, overlay);

            Assert.True(result.IsValid);
            Assert.Equal(42, Assert.Single(result.Rows).Size);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void DisabledColumnsAreIgnoredDuringImportEvenWhenTheyArePresent()
    {
        var path = TemporaryPath();
        var validationCalls = 0;
        var converter = new CountingConverter();
        var schema = Excel.Schema<Product>()
            .Column(product => product.Name)
            .Column(product => product.SerialNumber, column => column
                .ConvertWith(converter)
                .Validate(_ => { validationCalls++; return false; }, "must not run"))
            .Build();
        var overlay = Excel.Overlay<Product>(runtime => runtime.Include(product => product.SerialNumber, false));

        try
        {
            CreateWorkbook(path, ["Name", "SerialNumber"], ["Shirt", "SN-1"]);
            var result = Excel.Read(path, schema, overlay);

            Assert.True(result.IsValid);
            Assert.Equal("Shirt", Assert.Single(result.Rows).Name);
            Assert.Equal(0, converter.ReadCalls);
            Assert.Equal(0, validationCalls);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void DisabledColumnsDoNotWriteConvertersOrNativeValidation()
    {
        var path = TemporaryPath();
        var converter = new CountingConverter();
        var schema = Excel.Schema<Product>()
            .Column(product => product.Name)
            .Column(product => product.SerialNumber, column => column.ConvertWith(converter).AllowedValues("SN-1", "SN-2"))
            .Build();
        var overlay = Excel.Overlay<Product>(runtime => runtime.Include(product => product.SerialNumber, false));

        try
        {
            Excel.Write(path, [new Product { Name = "Shirt", SerialNumber = "SN-1" }], schema, overlay);
            using var document = SpreadsheetDocument.Open(path, false);

            Assert.Equal(["Name"], Headers(document));
            Assert.Equal(0, converter.WriteCalls);
            Assert.Null(document.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().Single().State);
            Assert.DoesNotContain(document.WorkbookPart.Workbook!.Sheets.Elements<Sheet>(), sheet => (sheet.Name?.Value ?? string.Empty).StartsWith("_CellSharpValidation", StringComparison.Ordinal));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void OverlayRejectsInvalidConfigurationBeforeIo()
    {
        var schema = Excel.Schema<Product>()
            .Column(product => product.Name)
            .Column(product => product.Size, column => column.MapFromHeader("Legacy size"))
            .Column(product => product.Ignored, column => column.Ignore())
            .Build();

        Assert.Throws<ArgumentException>(() => Excel.Write(TemporaryPath(), Array.Empty<Product>(), schema,
            Excel.Overlay<Product>(runtime => runtime.Header(product => product.Size, "Name"))));
        Assert.Throws<ArgumentException>(() => Excel.Write(TemporaryPath(), Array.Empty<Product>(), schema,
            Excel.Overlay<Product>(runtime => runtime.Header(product => product.Name, "Legacy size"))));
        Assert.Throws<ArgumentException>(() => Excel.Write(TemporaryPath(), Array.Empty<Product>(), schema,
            Excel.Overlay<Product>(runtime => runtime.Header(product => product.Ignored, "Ignored"))));
        Assert.Throws<ArgumentException>(() => Excel.Write(TemporaryPath(), Array.Empty<Product>(), schema,
            Excel.Overlay<Product>(runtime => runtime.Header(product => product.SerialNumber, "Serial"))));
        Assert.Throws<ArgumentException>(() => Excel.Overlay<Product>(runtime => runtime.Header(product => product.Name, " ")));
        Assert.Throws<InvalidOperationException>(() => Excel.Overlay<Product>(runtime => runtime
            .Header(product => product.Name, "One")
            .Header(product => product.Name, "Two")));
        Assert.Throws<InvalidOperationException>(() => Excel.Overlay<Product>(runtime => runtime
            .Include(product => product.Name, true)
            .Include(product => product.Name, false)));
    }

    [Fact]
    public void SchemaAndOverlaysCanBeReusedIndependently()
    {
        var italianPath = TemporaryPath();
        var englishPath = TemporaryPath();
        var schema = ProductSchema();
        var italian = Excel.Overlay<Product>(runtime => runtime.Header(product => product.Size, "Taglia"));
        var english = Excel.Overlay<Product>(runtime => runtime.Header(product => product.Size, "Size"));

        try
        {
            Excel.Write(italianPath, [new Product { Name = "Shirt", Size = 42 }], schema, italian);
            Excel.Write(englishPath, [new Product { Name = "Shirt", Size = 42 }], schema, english);

            Assert.Contains("Taglia", Headers(italianPath));
            Assert.Contains("Size", Headers(englishPath));
            Assert.Equal("Size", Headers(englishPath).Single(header => header == "Size"));
        }
        finally
        {
            Delete(italianPath);
            Delete(englishPath);
        }
    }

    [Fact]
    public void WorkbookUsesIndependentOverlaysForEachTypedWorksheet()
    {
        var path = TemporaryPath();
        var products = ProductSchema();
        var stock = Excel.Schema<Stock>()
            .SheetName("Stock")
            .Column(item => item.Quantity)
            .Build();
        var productOverlay = Excel.Overlay<Product>(runtime => runtime.Header(product => product.Name, "Product name"));
        var stockOverlay = Excel.Overlay<Stock>(runtime => runtime.Header(item => item.Quantity, "Available"));

        try
        {
            Excel.Workbook()
                .AddSheet([new Product { Name = "Shirt", Size = 42 }], products, productOverlay)
                .AddSheet([new Stock { Quantity = 3 }], stock, stockOverlay)
                .Write(path);

            using var workbook = Excel.Open(path);
            Assert.Equal("Shirt", Assert.Single(workbook.Read(products, productOverlay).Rows).Name);
            Assert.Equal(3, Assert.Single(workbook.Read(stock, stockOverlay).Rows).Quantity);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void MultiSheetTemplatesUseIndependentOverlays()
    {
        var path = TemporaryPath();
        var products = ProductSchema();
        var stock = Excel.Schema<Stock>()
            .SheetName("Stock")
            .Column(item => item.Quantity)
            .Build();
        var productOverlay = Excel.Overlay<Product>(runtime => runtime
            .Header(product => product.Size, "Taglia")
            .Include(product => product.SerialNumber, false));
        var stockOverlay = Excel.Overlay<Stock>(runtime => runtime.Header(item => item.Quantity, "Available"));

        try
        {
            Excel.Workbook()
                .AddTemplateSheet(products, productOverlay)
                .AddTemplateSheet(stock, stockOverlay)
                .CreateTemplate(path);

            using var document = SpreadsheetDocument.Open(path, false);
            Assert.Equal(["Name", "Taglia"], Headers(document, "Products"));
            Assert.Equal(["Available"], Headers(document, "Stock"));
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void PositionalMappingsRemainPhysicalWhenOtherColumnsAreDisabled()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<Product>()
            .Column(product => product.Name)
            .Column(product => product.Size, column => column.MapFromColumn(3))
            .Column(product => product.SerialNumber, column => column.Optional())
            .Build();
        var overlay = Excel.Overlay<Product>(runtime => runtime.Include(product => product.Name, false));

        try
        {
            CreateWorkbook(path, ["Unused", "Still unused", "Physical size"], ["x", "y", "42"]);
            var result = Excel.Read(path, schema, overlay);

            Assert.True(result.IsValid);
            Assert.Equal(42, Assert.Single(result.Rows).Size);
        }
        finally
        {
            Delete(path);
        }
    }

    private static ExcelSchema<Product> ProductSchema() => Excel.Schema<Product>()
        .SheetName("Products")
        .Column(product => product.Name)
        .Column(product => product.Size)
        .Column(product => product.SerialNumber, column => column.Optional())
        .Build();

    private static string[] Headers(string path)
    {
        using var document = SpreadsheetDocument.Open(path, false);
        return Headers(document);
    }

    private static string[] Headers(SpreadsheetDocument document) => document.WorkbookPart!.WorksheetParts.First().Worksheet!
        .GetFirstChild<SheetData>()!.Elements<Row>().First().Elements<Cell>()
        .Select(cell => cell.InlineString!.Text!.Text).ToArray();

    private static string[] Headers(SpreadsheetDocument document, string sheetName)
    {
        var sheet = document.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().Single(item => item.Name!.Value == sheetName);
        var worksheet = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id!);
        return worksheet.Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().First().Elements<Cell>()
            .Select(cell => cell.InlineString!.Text!.Text).ToArray();
    }

    private static void CreateWorkbook(string path, string[] headers, params string[][] rows)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbook = document.AddWorkbookPart();
        workbook.Workbook = new Workbook();
        var worksheet = workbook.AddNewPart<WorksheetPart>();
        var data = new SheetData(CreateRow(headers));
        foreach (var row in rows)
        {
            data.AppendChild(CreateRow(row));
        }

        worksheet.Worksheet = new Worksheet(data);
        var sheets = workbook.Workbook!.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet { Id = workbook.GetIdOfPart(worksheet), SheetId = 1U, Name = "Sheet1" });
        worksheet.Worksheet!.Save();
        workbook.Workbook!.Save();
    }

    private static Row CreateRow(IEnumerable<string> values) => new(values.Select(value => new Cell
    {
        DataType = CellValues.InlineString,
        InlineString = new InlineString(new Text(value)),
    }));

    private static string TemporaryPath() => Path.Combine(Path.GetTempPath(), $"cellsharp-runtime-schema-{Guid.NewGuid():N}.xlsx");

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
        public int Size { get; set; }
        public string? SerialNumber { get; set; }
        public string? Ignored { get; set; }
    }

    private sealed class Stock
    {
        public int Quantity { get; set; }
    }

    private sealed class CountingConverter : IExcelValueConverter<string?, string>
    {
        public int WriteCalls { get; private set; }
        public int ReadCalls { get; private set; }

        public string Write(string? value)
        {
            WriteCalls++;
            return value ?? string.Empty;
        }

        public bool TryRead(string value, out string? converted)
        {
            ReadCalls++;
            converted = value;
            return true;
        }
    }
}
