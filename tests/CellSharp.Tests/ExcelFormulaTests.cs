using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelFormulaTests
{
    [Fact]
    public void WriteEmitsNativeNormalizedFormulasWithCachedValuesStylesAndRecalculation()
    {
        var path = TemporaryPath();
        var schema = FormulaSchema();

        try
        {
            Excel.Write(path,
            [
                new Invoice { Quantity = 2, UnitPrice = 12.5M, Total = 999M },
                new Invoice { Quantity = 3, UnitPrice = 10M, Total = 999M },
                new Invoice { Quantity = 4, UnitPrice = 7.5M, Total = 999M },
            ], schema, options => options.AutoFitColumns());

            using var document = SpreadsheetDocument.Open(path, false);
            var worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            var totals = DataCells(worksheet).Select(row => row.ElementAt(2)).ToArray();
            var calculation = document.WorkbookPart.Workbook!.CalculationProperties;
            var totalColumn = worksheet.GetFirstChild<Columns>()!.Elements<Column>().Single(column => column.Min!.Value == 3U);
            var totalFormat = document.WorkbookPart.WorkbookStylesPart!.Stylesheet!.CellFormats!
                .Elements<CellFormat>()
                .ElementAt((int)totals[0].StyleIndex!.Value);
            var totalNumberFormat = document.WorkbookPart.WorkbookStylesPart.Stylesheet!.NumberingFormats!
                .Elements<NumberingFormat>()
                .Single(format => format.NumberFormatId!.Value == totalFormat.NumberFormatId!.Value);

            Assert.Equal(["A2*B2", "A3*B3", "A4*B4"], totals.Select(total => total.CellFormula!.InnerText));
            Assert.Equal(["999", "999", "999"], totals.Select(total => total.CellValue!.Text));
            Assert.All(totals, total =>
            {
                Assert.Equal(CellValues.Number, total.DataType!.Value);
                Assert.NotNull(total.StyleIndex);
            });
            Assert.Equal(18D, totalColumn.Width!.Value);
            Assert.Equal(HorizontalAlignmentValues.Right, totalFormat.Alignment!.Horizontal!.Value);
            Assert.Equal("#,##0.00", totalNumberFormat.FormatCode!.Value);
            Assert.True(calculation!.FullCalculationOnLoad!.Value);
            Assert.True(calculation.ForceFullCalculation!.Value);
            Assert.True(calculation.CalculationOnSave!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void FormulaExportCanBeImmediatelyImportedUsingCachedPropertyValues()
    {
        var path = TemporaryPath();
        var schema = FormulaSchema();

        try
        {
            Excel.Write(path,
            [
                new Invoice { Quantity = 2, UnitPrice = 12.5M, Total = 25M },
                new Invoice { Quantity = 3, UnitPrice = 10M, Total = 30M },
            ], schema);

            var result = Excel.Read<Invoice>(path, schema);

            Assert.True(result.IsValid);
            Assert.Equal([25M, 30M], result.Rows.Select(row => row.Total));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void FormulaContextUsesOneBasedExcelRowsAndPreservesCrossSheetReferences()
    {
        var path = TemporaryPath();
        var sourceSchema = Excel.Schema<SourceRow>()
            .SheetName("Source")
            .Column(row => row.Amount)
            .Build();
        var reportSchema = Excel.Schema<ReportRow>()
            .SheetName("Report")
            .Column(row => row.Total, column => column.Formula(context => $"=Source!A{context.Row}"))
            .Build();

        try
        {
            Excel.Workbook()
                .AddSheet([new SourceRow { Amount = 10M }, new SourceRow { Amount = 20M }], sourceSchema)
                .AddSheet([new ReportRow(), new ReportRow()], reportSchema)
                .Write(path);

            using var document = SpreadsheetDocument.Open(path, false);
            var report = Worksheet(document, "Report");
            var formulas = DataCells(report).Select(row => Assert.Single(row).CellFormula!.InnerText).ToArray();

            Assert.Equal(["Source!A2", "Source!A3"], formulas);
            Assert.True(document.WorkbookPart!.Workbook!.CalculationProperties!.FullCalculationOnLoad!.Value);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void TemplatesKeepFormulaColumnsHeaderOnlyAndDoNotRequestRecalculation()
    {
        var path = TemporaryPath();

        try
        {
            Excel.CreateTemplate(path, FormulaSchema());

            using var document = SpreadsheetDocument.Open(path, false);
            var worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet!;

            Assert.Single(worksheet.GetFirstChild<SheetData>()!.Elements<Row>());
            Assert.Null(document.WorkbookPart.Workbook!.CalculationProperties);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void EmptyFormulaExportDoesNotRequestRecalculationBeforeAnyFormulaCellExists()
    {
        var path = TemporaryPath();

        try
        {
            Excel.Write(path, Array.Empty<Invoice>(), FormulaSchema());

            using var document = SpreadsheetDocument.Open(path, false);

            Assert.Null(document.WorkbookPart!.Workbook!.CalculationProperties);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void FormulaCachedValuesUseTheNormalImportAndConverterPipeline()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<ConvertedInvoice>()
            .Column(invoice => invoice.Total, column => column
                .Formula(context => $"A{context.Row}*B{context.Row}")
                .ConvertWith(IncrementingIntConverter.Instance))
            .Build();

        try
        {
            CreateFormulaWorkbook(path, "Total", "A2*B2", "7");

            var result = Excel.Read<ConvertedInvoice>(path, schema);

            Assert.True(result.IsValid);
            Assert.Equal(8, Assert.Single(result.Rows).Total);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void FormulaWithoutCachedValueIsInvalidAndIncludeRetainsThePartialRow()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<Invoice>()
            .Column(invoice => invoice.Total, column => column.Formula(context => $"A{context.Row}*B{context.Row}"))
            .Build();

        try
        {
            CreateFormulaWorkbook(path, "Total", "A2*B2", null);

            var skipped = Excel.Read<Invoice>(path, schema);
            var included = Excel.Read(path, schema, new ExcelReadOptions(ExcelInvalidRowPolicy.Include));

            Assert.Empty(skipped.Rows);
            Assert.Equal(ExcelReadErrorCode.InvalidValue, Assert.Single(skipped.Errors).Code);
            Assert.False(included.IsValid);
            Assert.Equal(0M, Assert.Single(included.Rows).Total);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void FormulaExportUsesThePropertyValueAsItsCachedResult()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<ConvertedInvoice>()
            .Column(invoice => invoice.Total, column => column
                .ConvertWith(IncrementingIntConverter.Instance)
                .Formula(context => $"=1+{context.Row}"))
            .Build();

        try
        {
            Excel.Write(path, [new ConvertedInvoice { Total = 42 }], schema);

            using var document = SpreadsheetDocument.Open(path, false);
            var cell = Assert.Single(DataCells(document.WorkbookPart!.WorksheetParts.Single().Worksheet!)).Single();
            var formula = cell.CellFormula;

            Assert.Equal("1+2", formula!.InnerText);
            Assert.Equal("42", cell.CellValue!.Text);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void FormulaConfigurationRejectsIncompatibleCombinationsAndInvalidExpressions()
    {
        Assert.Throws<InvalidOperationException>(() => Excel.Schema<FormulaText>()
            .Column(row => row.Value, column => column.Formula(_ => "A2").AllowedValues("A"))
            .Build());
        Assert.Throws<InvalidOperationException>(() => Excel.Schema<FormulaText>()
            .Column(row => row.Value, column => column.Formula(_ => "A2").Ignore())
            .Build());
        Assert.Throws<InvalidOperationException>(() => Excel.Schema<Invoice>()
            .Column(row => row.Total, column => column.Formula(_ => "A2").Range(0M, 10M))
            .Build());
        Assert.Throws<InvalidOperationException>(() => Excel.Schema<FormulaDate>()
            .Column(row => row.Value, column => column.Formula(_ => "A2").DateBetween(DateTime.UnixEpoch, DateTime.UnixEpoch.AddDays(1)))
            .Build());

        var emptyPath = TemporaryPath();
        var doubledPath = TemporaryPath();
        try
        {
            var empty = Excel.Schema<FormulaText>()
                .Column(row => row.Value, column => column.Formula(_ => " "))
                .Build();
            var doubled = Excel.Schema<FormulaText>()
                .Column(row => row.Value, column => column.Formula(_ => "==A2"))
                .Build();

            Assert.Throws<InvalidOperationException>(() => Excel.Write(emptyPath, [new FormulaText()], empty));
            Assert.Throws<InvalidOperationException>(() => Excel.Write(doubledPath, [new FormulaText()], doubled));
        }
        finally
        {
            Delete(emptyPath);
            Delete(doubledPath);
        }
    }

    [Fact]
    public void FormulaExportStillObservesCooperativeCancellationBetweenRows()
    {
        using var stream = new MemoryStream();
        using var source = new CancellationTokenSource();

        Assert.Throws<OperationCanceledException>(() => Excel.Write(stream, RowsThatCancel(source), FormulaSchema(), source.Token));
    }

    private static ExcelSchema<Invoice> FormulaSchema() => Excel.Schema<Invoice>()
        .Column(invoice => invoice.Quantity)
        .Column(invoice => invoice.UnitPrice)
        .Column(invoice => invoice.Total, column => column
            .Formula(context => $"=A{context.Row}*B{context.Row}")
            .Format("#,##0.00")
            .Width(18D)
            .Align(ExcelHorizontalAlignment.Right))
        .Build();

    private static IEnumerable<Cell[]> DataCells(Worksheet worksheet) => worksheet.GetFirstChild<SheetData>()!
        .Elements<Row>()
        .Skip(1)
        .Select(row => row.Elements<Cell>().ToArray());

    private static Worksheet Worksheet(SpreadsheetDocument document, string name)
    {
        var sheet = document.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().Single(sheet => sheet.Name!.Value == name);
        return ((WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id!)).Worksheet!;
    }

    private static void CreateFormulaWorkbook(string path, string header, string formula, string? cachedValue)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var formulaCell = new Cell { CellReference = "A2", CellFormula = new CellFormula(formula) };
        if (cachedValue is not null)
        {
            formulaCell.CellValue = new CellValue(cachedValue);
        }

        worksheetPart.Worksheet = new Worksheet(new SheetData(
            new Row(new Cell
            {
                CellReference = "A1",
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new Text(header)),
            }),
            new Row(formulaCell)));
        var sheets = workbookPart.Workbook!.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1U, Name = "Sheet1" });
        worksheetPart.Worksheet!.Save();
        workbookPart.Workbook!.Save();
    }

    private static IEnumerable<Invoice> RowsThatCancel(CancellationTokenSource source)
    {
        yield return new Invoice { Quantity = 2, UnitPrice = 12.5M };
        source.Cancel();
        yield return new Invoice { Quantity = 3, UnitPrice = 10M };
    }

    private static string TemporaryPath() => Path.Combine(Path.GetTempPath(), $"cellsharp-formula-{Guid.NewGuid():N}.xlsx");

    private static void Delete(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }

    private sealed class Invoice
    {
        public int Quantity { get; set; }

        public decimal UnitPrice { get; set; }

        public decimal Total { get; set; }
    }

    private sealed class SourceRow
    {
        public decimal Amount { get; set; }
    }

    private sealed class ReportRow
    {
        public decimal Total { get; set; }
    }

    private sealed class ConvertedInvoice
    {
        public int Total { get; set; }
    }

    private sealed class FormulaText
    {
        public string? Value { get; set; }
    }

    private sealed class FormulaDate
    {
        public DateTime Value { get; set; }
    }

    private sealed class IncrementingIntConverter : IExcelValueConverter<int, int>
    {
        internal static IncrementingIntConverter Instance { get; } = new();

        public int Write(int value) => value;

        public bool TryRead(int value, out int converted)
        {
            converted = value + 1;
            return true;
        }
    }

}
