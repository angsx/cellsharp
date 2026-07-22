using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelCancellationTests
{
    [Fact]
    public void AlreadyCancelledTokensFailBeforeReadWriteAndTemplateWithoutClosingTheStream()
    {
        using var stream = new TrackingStream();
        var schema = Schema();
        using var source = new CancellationTokenSource();
        source.Cancel();

        Assert.Throws<OperationCanceledException>(() => Excel.Write(stream, Array.Empty<Row>(), source.Token));
        Assert.Throws<OperationCanceledException>(() => Excel.CreateTemplate(stream, schema, source.Token));
        Assert.Throws<OperationCanceledException>(() => Excel.Read<Row>(stream, source.Token));
        Assert.False(stream.WasDisposed);
    }

    [Fact]
    public void ExportCancellationIsObservedBetweenRowsAndLeavesTheCallerStreamOpen()
    {
        using var stream = new TrackingStream();
        using var source = new CancellationTokenSource();

        Assert.Throws<OperationCanceledException>(() => Excel.Write(stream, RowsThatCancel(source), source.Token));

        Assert.False(stream.WasDisposed);
    }

    [Fact]
    public void ImportCancellationIsObservedBetweenRowsAndIsNotAnImportDiagnostic()
    {
        using var stream = new MemoryStream();
        using var source = new CancellationTokenSource();
        var schema = Excel.Schema<Row>()
            .Column(row => row.Name)
            .Column(row => row.Value, column => column.Validate(value =>
            {
                source.Cancel();
                return true;
            }, "unused"))
            .Build();

        Excel.Write(stream, [new Row { Name = "A", Value = 1 }, new Row { Name = "B", Value = 2 }], schema);

        Assert.Throws<OperationCanceledException>(() => Excel.Read(stream, schema, source.Token));
    }

    [Fact]
    public void MultiSheetCancellationStopsTheWorkbookBetweenSheetsAndLeavesTheStreamOpen()
    {
        using var stream = new TrackingStream();
        using var source = new CancellationTokenSource();
        var first = Excel.Schema<Row>().SheetName("First").Column(row => row.Name).Build();
        var second = Excel.Schema<Row>().SheetName("Second").Column(row => row.Name).Build();

        Assert.Throws<OperationCanceledException>(() => Excel.Workbook()
            .AddSheet(RowsThatCancel(source), first)
            .AddSheet([new Row { Name = "Second" }], second)
            .Write(stream, source.Token));

        Assert.False(stream.WasDisposed);
    }

    [Fact]
    public void ConverterCancellationPropagatesInsteadOfBecomingAnImportDiagnostic()
    {
        using var stream = new MemoryStream();
        var schema = Excel.Schema<Row>()
            .Column(row => row.Value, column => column.ConvertWith(new CancellingConverter()))
            .Build();

        Excel.Write(stream, [new Row { Value = 1 }], schema);

        Assert.Throws<OperationCanceledException>(() => Excel.Read(stream, schema));
    }

    private static ExcelSchema<Row> Schema() => Excel.Schema<Row>()
        .Column(row => row.Name)
        .Column(row => row.Value)
        .Build();

    private static IEnumerable<Row> RowsThatCancel(CancellationTokenSource source)
    {
        yield return new Row { Name = "First", Value = 1 };
        source.Cancel();
        yield return new Row { Name = "Second", Value = 2 };
    }

    private sealed class Row
    {
        public string? Name { get; set; }
        public int Value { get; set; }
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

    private sealed class CancellingConverter : IExcelValueConverter<int, int>
    {
        public int Write(int value) => value;

        public bool TryRead(int value, out int converted)
        {
            throw new OperationCanceledException();
        }
    }
}
