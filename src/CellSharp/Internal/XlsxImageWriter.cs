using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CellSharp.Internal;

internal static class XlsxImageWriter
{
    private const long EmusPerPixel = 9525L; // 96 DPI

    internal static void Apply(IReadOnlyList<WorksheetValidationContext> worksheets)
    {
        foreach (var context in worksheets)
        {
            var images = context.Layout?.Images;
            if (images is null || images.Count == 0) continue;
            var drawingsPart = context.WorksheetPart.AddNewPart<DrawingsPart>();
            var drawing = new Xdr.WorksheetDrawing();
            drawingsPart.WorksheetDrawing = drawing;
            for (var index = 0; index < images.Count; index++) AddImage(drawingsPart, drawing, images[index], (uint)(index + 1));
            context.Worksheet.AppendChild(new Drawing { Id = context.WorksheetPart.GetIdOfPart(drawingsPart) });
        }
    }

    private static void AddImage(DrawingsPart drawingsPart, Xdr.WorksheetDrawing drawing, WorksheetImageDefinition image, uint pictureId)
    {
        var (width, height) = image.ResolveSize();
        var imagePart = drawingsPart.AddImagePart(image.Format == ExcelImageFormat.Png ? ImagePartType.Png : ImagePartType.Jpeg);
        using (var stream = new MemoryStream(image.Data, writable: false)) imagePart.FeedData(stream);
        var relationshipId = drawingsPart.GetIdOfPart(imagePart);
        var anchor = image.Anchor!.Value;
        var from = new Xdr.FromMarker(
            new Xdr.ColumnId((anchor.FromColumn - 1).ToString(System.Globalization.CultureInfo.InvariantCulture)),
            new Xdr.ColumnOffset(ToEmu(image.OffsetX).ToString(System.Globalization.CultureInfo.InvariantCulture)),
            new Xdr.RowId((anchor.FromRow - 1).ToString(System.Globalization.CultureInfo.InvariantCulture)),
            new Xdr.RowOffset(ToEmu(image.OffsetY).ToString(System.Globalization.CultureInfo.InvariantCulture)));
        var extent = new Xdr.Extent { Cx = ToEmu(width), Cy = ToEmu(height) };
        var picture = new Xdr.Picture(
            new Xdr.NonVisualPictureProperties(
                new Xdr.NonVisualDrawingProperties { Id = pictureId, Name = image.Name ?? $"Image {pictureId}", Description = image.AltText },
                new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })),
            new Xdr.BlipFill(
                new A.Blip { Embed = relationshipId, CompressionState = A.BlipCompressionValues.Print },
                new A.Stretch(new A.FillRectangle())),
            new Xdr.ShapeProperties(
                new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = ToEmu(width), Cy = ToEmu(height) }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));
        drawing.AppendChild(new Xdr.OneCellAnchor(from, extent, picture, new Xdr.ClientData()));
    }

    private static long ToEmu(int pixels) => checked(pixels * EmusPerPixel);
}
