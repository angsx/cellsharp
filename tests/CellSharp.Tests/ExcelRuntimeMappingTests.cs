using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelRuntimeMappingTests
{
    [Fact]
    public void MapFromColumnUsesOneBasedWorksheetColumnNumbers()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<Product>()
            .Column(product => product.Color, column => column.MapFromColumn(3))
            .Build();

        try
        {
            CreateWorkbook(path, ["Unused", "Also unused", "Runtime value"], ["a", "b", "Blue"]);

            var result = Excel.Read<Product>(path, schema);

            Assert.True(result.IsValid);
            Assert.Equal("Blue", Assert.Single(result.Rows).Color);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void MapFromHeaderUsesTheExistingConversionAndValidationPipeline()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<Product>()
            .Column(product => product.Size, column => column
                .MapFromHeader("size")
                .Range(1, 10))
            .Build();

        try
        {
            CreateWorkbook(path, ["SIZE"], ["7"], ["11"]);

            var result = Excel.Read<Product>(path, schema);

            Assert.Equal(7, Assert.Single(result.Rows).Size);
            var error = Assert.Single(result.Errors);
            Assert.Equal(ExcelReadErrorCode.ValidationFailed, error.Code);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void RuntimeMappingsRespectOptionalColumnsAndCustomConverters()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<Product>()
            .Column(product => product.OptionalColor, column => column
                .Optional()
                .MapFromHeader("Colour"))
            .Column(product => product.Status, column => column
                .MapFromHeader("State")
                .ConvertWith(StatusConverter.Instance)
                .Validate(status => status != Status.Archived, "Archived products are not importable"))
            .Build();

        try
        {
            CreateWorkbook(path, ["State"], ["ACTIVE"], ["ARCHIVED"]);

            var result = Excel.Read<Product>(path, schema);

            Assert.Null(Assert.Single(result.Rows).OptionalColor);
            Assert.Equal(Status.Active, result.Rows[0].Status);
            var error = Assert.Single(result.Errors);
            Assert.Equal(ExcelReadErrorCode.ValidationFailed, error.Code);
            Assert.Equal("Archived products are not importable", error.Message);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void PositionalMappingsTakePrecedenceOverRuntimeAndSchemaHeaders()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<Product>()
            .Column(product => product.Color, column => column
                .Header("Colour")
                .MapFromColumn(3)
                .MapFromHeader("Color Alias"))
            .Build();

        try
        {
            CreateWorkbook(path, ["Unused", "Also unused", "Colour", "Color Alias"], ["a", "b", "Positional", "Header"]);

            var result = Excel.Read<Product>(path, schema);

            Assert.True(result.IsValid);
            Assert.Equal("Positional", Assert.Single(result.Rows).Color);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void MissingRequiredRuntimeMappingsUseTheExistingMissingHeaderError()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<Product>()
            .Column(product => product.Color, column => column.MapFromColumn(3))
            .Build();

        try
        {
            CreateWorkbook(path, ["First", "Second"], ["a", "b"]);

            var result = Excel.Read<Product>(path, schema);

            var error = Assert.Single(result.Errors);
            Assert.Equal(ExcelReadErrorCode.MissingHeader, error.Code);
            Assert.Equal("Color", error.Header);
            Assert.Contains("Column 3", error.Message);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void RuntimeMappingConflictsAreRejectedWhenTheSchemaIsBuilt()
    {
        Assert.Throws<ArgumentException>(() => Excel.Schema<Product>()
            .Column(product => product.Color, column => column.MapFromColumn(3))
            .Column(product => product.Size, column => column.MapFromColumn(3))
            .Build());

        Assert.Throws<ArgumentException>(() => Excel.Schema<Product>()
            .Column(product => product.Color, column => column.MapFromHeader("Color"))
            .Column(product => product.Size, column => column.MapFromHeader("color"))
            .Build());

        Assert.Throws<ArgumentException>(() => Excel.Schema<Product>()
            .Column(product => product.Color, column => column.MapFromHeader("Size"))
            .Column(product => product.Size)
            .Build());
    }

    [Fact]
    public void RuntimeMappingRejectsInvalidColumnNumbersAndIgnoredColumns()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => Excel.Schema<Product>()
            .Column(product => product.Color, column => column.MapFromColumn(0)));

        Assert.Throws<InvalidOperationException>(() => Excel.Schema<Product>()
            .Column(product => product.Color, column => column.MapFromHeader("Color").Ignore()));
    }

    private static void CreateWorkbook(string path, string[] headers, params string[][] values)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var data = new SheetData(Row(headers));
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

    private static string TemporaryPath() => Path.Combine(Path.GetTempPath(), $"cellsharp-runtime-mapping-{Guid.NewGuid():N}.xlsx");

    private static void Delete(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }

    private sealed class Product
    {
        public string? Color { get; set; }

        public string? OptionalColor { get; set; }

        public int Size { get; set; }

        public Status? Status { get; set; }
    }

    private sealed class Status : IEquatable<Status>
    {
        internal static Status Active { get; } = new("ACTIVE");
        internal static Status Archived { get; } = new("ARCHIVED");

        internal Status(string value) => Value = value;

        internal string Value { get; }

        public bool Equals(Status? other) => other is not null && Value == other.Value;

        public override bool Equals(object? obj) => Equals(obj as Status);

        public override int GetHashCode() => Value.GetHashCode();
    }

    private sealed class StatusConverter : IExcelValueConverter<Status?, string>
    {
        internal static StatusConverter Instance { get; } = new();

        public string Write(Status? value) => value?.Value ?? string.Empty;

        public bool TryRead(string value, out Status? converted)
        {
            converted = value == "ACTIVE" ? Status.Active : value == "ARCHIVED" ? Status.Archived : null;
            return converted is not null;
        }
    }
}
