using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelStreamTests
{
    [Fact]
    public void SingleSheetStreamRoundTripSupportsSchemaOverlayOptionsAndCallerOwnership()
    {
        using var stream = new TrackingStream();
        var schema = Excel.Schema<Product>()
            .Column(product => product.Name)
            .Column(product => product.Quantity, column => column.Range(1, 10))
            .Column(product => product.SerialNumber, column => column.Optional())
            .Build();
        var overlay = Excel.Overlay<Product>(runtime => runtime
            .Header(product => product.Name, "Product")
            .Include(product => product.SerialNumber, false));

        Excel.Write(stream, [new Product { Name = "Shirt", Quantity = 3, SerialNumber = "SN" }], schema, overlay);

        Assert.False(stream.WasDisposed);
        Assert.Equal(stream.Length, stream.Position);
        stream.Position = 19;
        var result = Excel.Read(stream, schema, overlay, new ExcelReadOptions(ExcelInvalidRowPolicy.Include));

        Assert.True(result.IsValid);
        Assert.Equal("Shirt", Assert.Single(result.Rows).Name);
        Assert.False(stream.WasDisposed);
    }

    [Fact]
    public void StreamTemplatePreservesPresentationAndOwnership()
    {
        using var stream = new TrackingStream();
        var schema = Excel.Schema<Product>()
            .Column(product => product.Name, column => column.Width(18D))
            .Column(product => product.Quantity, column => column.Header("Count"))
            .Build();
        var overlay = Excel.Overlay<Product>(runtime => runtime.Header(product => product.Quantity, "Available"));

        Excel.CreateTemplate(stream, schema, overlay, options => options.FreezeHeaderRow());

        Assert.False(stream.WasDisposed);
        Assert.Equal(stream.Length, stream.Position);
        stream.Position = 0;
        using var document = SpreadsheetDocument.Open(stream, false);
        var row = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().Single();
        Assert.Equal(["Name", "Available"], row.Elements<Cell>().Select(Text));
    }

    [Fact]
    public void MultiSheetStreamUsesOneHiddenValidationLookupAndOpenLeavesStreamOpen()
    {
        using var stream = new TrackingStream();
        var customers = Excel.Schema<Product>()
            .SheetName("Customers")
            .Column(product => product.Name)
            .Column(product => product.Category, column => column.AllowedValues("Retail", "Wholesale"))
            .Build();
        var orders = Excel.Schema<Order>()
            .SheetName("Orders")
            .Column(order => order.Id)
            .Column(order => order.Category, column => column.AllowedValues("Retail", "Wholesale"))
            .Build();

        Excel.Workbook()
            .AddSheet([new Product { Name = "Ada", Category = "Retail" }], customers)
            .AddSheet([new Order { Id = 7, Category = "Wholesale" }], orders)
            .Write(stream);

        Assert.False(stream.WasDisposed);
        stream.Position = 0;
        using (var workbook = Excel.Open(stream))
        {
            Assert.Equal("Ada", Assert.Single(workbook.Read(customers).Rows).Name);
            Assert.Equal(7, Assert.Single(workbook.ReadAt(1, orders).Rows).Id);
        }

        Assert.False(stream.WasDisposed);
        stream.Position = 0;
        using var document = SpreadsheetDocument.Open(stream, false);
        Assert.Single(document.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>(), sheet => sheet.State?.Value == SheetStateValues.Hidden);
    }

    [Fact]
    public void StreamWritesTruncateExistingContentAndReplaceWorkbook()
    {
        using var stream = new TrackingStream();
        Excel.Write(stream, Enumerable.Range(1, 100).Select(number => new Product { Name = new string('X', 100), Quantity = number }));
        var largeLength = stream.Length;

        Excel.Write(stream, [new Product { Name = "Small", Quantity = 1 }]);

        Assert.True(stream.Length < largeLength);
        Assert.Equal(stream.Length, stream.Position);
        stream.Position = 0;
        Assert.Equal("Small", Assert.Single(Excel.Read<Product>(stream).Rows).Name);
    }

    [Fact]
    public void StreamApisRejectUnsupportedCapabilitiesWithoutClosingTheCallerStream()
    {
        using var readable = new TrackingStream();
        using var nonSeekable = new NonSeekableStream(readable);
        using var closed = new MemoryStream();
        closed.Dispose();

        Assert.Throws<NotSupportedException>(() => Excel.Write(nonSeekable, Array.Empty<Product>()));
        Assert.Throws<NotSupportedException>(() => Excel.Read<Product>(nonSeekable));
        Assert.Throws<ArgumentException>(() => Excel.Write(closed, Array.Empty<Product>()));
        Assert.Throws<ArgumentException>(() => Excel.Read<Product>(closed));
        Assert.False(readable.WasDisposed);
    }

    private static string Text(Cell cell) => cell.InlineString!.Text!.Text;

    private sealed class Product
    {
        public string? Name { get; set; }
        public int Quantity { get; set; }
        public string? SerialNumber { get; set; }
        public string? Category { get; set; }
    }

    private sealed class Order
    {
        public int Id { get; set; }
        public string? Category { get; set; }
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

    private sealed class NonSeekableStream : Stream
    {
        private readonly Stream _inner;

        internal NonSeekableStream(Stream inner) => _inner = inner;

        public override bool CanRead => _inner.CanRead;
        public override bool CanSeek => false;
        public override bool CanWrite => _inner.CanWrite;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() => _inner.Flush();
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);
    }
}
