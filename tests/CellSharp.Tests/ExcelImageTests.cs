using System.Buffers.Binary;
using A = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelImageTests
{
    [Fact]
    public void PngPathAndJpegStreamCreateNativeDrawingPartsAndAnchors()
    {
        var workbookPath = TemporaryPath();
        var imagePath = Path.Combine(Path.GetTempPath(), $"cellsharp-image-{Guid.NewGuid():N}.png");
        var jpegPath = Path.Combine(Path.GetTempPath(), $"cellsharp-image-{Guid.NewGuid():N}.jpg");
        File.WriteAllBytes(imagePath, Png(200, 100));
        File.WriteAllBytes(jpegPath, Jpeg());
        try
        {
            using var jpeg = new MemoryStream(Jpeg());
            Excel.Workbook().AddSheet("Images", sheet =>
            {
                sheet.AddImage(imagePath).At("A1");
                sheet.AddImage(jpeg, ExcelImageFormat.Jpeg).At(3, 2, offsetX: 8, offsetY: 6).Width(160).KeepAspectRatio().Name("Logo").AltText("Company logo");
                sheet.AddImage(jpegPath).At("D1").Size(20, 20);
            }).Write(workbookPath);

            using var document = SpreadsheetDocument.Open(workbookPath, false);
            var part = document.WorkbookPart!.WorksheetParts.Single();
            var drawing = part.DrawingsPart!;
            Assert.NotNull(drawing);
            Assert.Equal(3, drawing.ImageParts.Count());
            var anchors = drawing.WorksheetDrawing!.Elements<A.OneCellAnchor>().ToArray();
            Assert.Equal(3, anchors.Length);
            Assert.Equal("0", anchors[0].FromMarker!.ColumnId!.Text);
            Assert.Equal(200L * 9525L, anchors[0].Extent!.Cx!.Value);
            Assert.Equal("1", anchors[1].FromMarker!.ColumnId!.Text);
            Assert.Equal(8L * 9525L, anchors[1].FromMarker!.ColumnOffset!.Text is { } offset ? long.Parse(offset, System.Globalization.CultureInfo.InvariantCulture) : 0L);
            Assert.Equal(160L * 9525L, anchors[1].Extent!.Cx!.Value);
            AssertValid(document);
        }
        finally { Delete(workbookPath); Delete(imagePath); Delete(jpegPath); }
    }

    [Fact]
    public void StreamIsSnapshottedWithoutClosingAndNonSeekableStreamsWork()
    {
        var path = TemporaryPath();
        var source = new NonSeekableReadStream(Png(120, 60));
        try
        {
            Excel.Workbook().AddSheet("Images", sheet => sheet.AddImage(source, ExcelImageFormat.Png).At("B2").Height(30).KeepAspectRatio()).Write(path);
            Assert.True(source.CanRead);
            using var document = SpreadsheetDocument.Open(path, false);
            var anchor = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing!.GetFirstChild<A.OneCellAnchor>()!;
            Assert.Equal(60L * 9525L, anchor.Extent!.Cx!.Value);
            Assert.Equal(30L * 9525L, anchor.Extent.Cy!.Value);
            AssertValid(document);
        }
        finally { source.Dispose(); Delete(path); }
    }

    [Fact]
    public void ExplicitSizeWinsAndImagesCoexistWithWorksheetFeaturesAndMultipleSheets()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<Entry>().SheetName("Report").AsTable()
            .Column(entry => entry.Status, column => column.AllowedValues("OK", "Review"))
            .Column(entry => entry.Amount)
            .Build();
        try
        {
            Excel.Workbook()
                .AddSheet([new Entry("OK", 1)], schema, sheet =>
                {
                    sheet.DataStartAt("B4");
                    sheet.Merge("A1:B1").Value("Report");
                    sheet.AddImage(new MemoryStream(Png(200, 100)), ExcelImageFormat.Png).At("A1").Size(160, 80).KeepAspectRatio();
                    sheet.Cell("C3").Comment("Image nearby").Hyperlink("https://openai.com");
                    sheet.Range("C5:C100").Name("Amounts").ConditionalFormat().GreaterThan(0);
                    sheet.Header.Center("Report");
                    sheet.Row(10).PageBreakAfter();
                })
                .AddSheet("Other", sheet => sheet.AddImage(new MemoryStream(Png(20, 10)), ExcelImageFormat.Png).At("C4"))
                .AddSheet("Empty", _ => { })
                .Write(path);

            using var document = SpreadsheetDocument.Open(path, false);
            var parts = document.WorkbookPart!.WorksheetParts.ToArray();
            Assert.NotNull(parts[0].DrawingsPart);
            Assert.NotNull(parts[1].DrawingsPart);
            Assert.Null(parts[2].DrawingsPart);
            var anchor = parts[0].DrawingsPart!.WorksheetDrawing!.GetFirstChild<A.OneCellAnchor>()!;
            Assert.Equal(160L * 9525L, anchor.Extent!.Cx!.Value);
            Assert.Equal(80L * 9525L, anchor.Extent.Cy!.Value);
            AssertValid(document);
        }
        finally { Delete(path); }
    }

    [Fact]
    public void InvalidImageInputsAndSizingFailFast()
    {
        Assert.Throws<FileNotFoundException>(() => Excel.Workbook().AddSheet("Images", sheet => sheet.AddImage("missing.png")));
        Assert.Throws<NotSupportedException>(() => Excel.Workbook().AddSheet("Images", sheet => sheet.AddImage("image.gif")));
        Assert.Throws<ArgumentException>(() => Excel.Workbook().AddSheet("Images", sheet => sheet.AddImage(new MemoryStream(), ExcelImageFormat.Png)));
        Assert.Throws<ArgumentException>(() => Excel.Workbook().AddSheet("Images", sheet => sheet.AddImage(new MemoryStream(Png(1, 1)), ExcelImageFormat.Jpeg)));
        Assert.Throws<ArgumentOutOfRangeException>(() => Excel.Workbook().AddSheet("Images", sheet => sheet.AddImage(new MemoryStream(Png(1, 1)), ExcelImageFormat.Png).At("A1").Width(0)));
        Assert.Throws<InvalidOperationException>(() => Excel.Workbook().AddSheet("Images", sheet => sheet.AddImage(new MemoryStream(Png(1, 1)), ExcelImageFormat.Png).At("A1").Width(10).Height(10).KeepAspectRatio()).Write(TemporaryPath()));
    }

    private static byte[] Png(int width, int height)
    {
        var bytes = new byte[24];
        new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }.CopyTo(bytes, 0);
        BinaryPrimitives.WriteInt32BigEndian(bytes.AsSpan(16, 4), width);
        BinaryPrimitives.WriteInt32BigEndian(bytes.AsSpan(20, 4), height);
        return bytes;
    }
    private static byte[] Jpeg() => Convert.FromBase64String("/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAP//////////////////////////////////////////////////////////////////////////////////////2wBDAf//////////////////////////////////////////////////////////////////////////////////////wAARCAABAAEDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAX/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/9oADAMBAAIQAxAAAAF//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABBQJ//8QAFBEBAAAAAAAAAAAAAAAAAAAAAP/aAAgBAwEBPwF//8QAFBEBAAAAAAAAAAAAAAAAAAAAAP/aAAgBAgEBPwF//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQAGPwJ//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABPyF//9k=");
    private static string TemporaryPath() => Path.Combine(Path.GetTempPath(), $"cellsharp-image-{Guid.NewGuid():N}.xlsx");
    private static void Delete(string path) { if (File.Exists(path)) File.Delete(path); }
    private static void AssertValid(SpreadsheetDocument document)
    {
        var errors = new OpenXmlValidator().Validate(document).ToArray();
        Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors.Select(error => error.Description)));
    }
    private sealed record Entry(string Status, int Amount);

    private sealed class NonSeekableReadStream : Stream
    {
        private readonly MemoryStream _inner;
        internal NonSeekableReadStream(byte[] data) => _inner = new MemoryStream(data);
        public override bool CanRead => _inner.CanRead; public override bool CanSeek => false; public override bool CanWrite => false; public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() { } public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count); public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException(); public override void SetLength(long value) => throw new NotSupportedException(); public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) { if (disposing) _inner.Dispose(); base.Dispose(disposing); }
    }
}
