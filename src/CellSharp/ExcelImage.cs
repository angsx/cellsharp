using CellSharp.Internal;

namespace CellSharp;

/// <summary>Image formats supported by CellSharp worksheet drawings.</summary>
public enum ExcelImageFormat
{
    /// <summary>Portable Network Graphics.</summary>
    Png,
    /// <summary>JPEG image data.</summary>
    Jpeg,
}

/// <summary>Configures one native worksheet image using pixel coordinates and dimensions.</summary>
public sealed class ExcelImage
{
    private readonly WorksheetImageDefinition _definition;
    internal ExcelImage(WorksheetImageDefinition definition) => _definition = definition;

    /// <summary>Anchors the image at an A1 cell reference with optional pixel offsets.</summary>
    public ExcelImage At(string reference, int offsetX = 0, int offsetY = 0)
    { _definition.SetAnchor(ExcelRangeReference.ParseCell(reference), offsetX, offsetY); return this; }
    /// <summary>Anchors the image at one-based worksheet coordinates with optional pixel offsets.</summary>
    public ExcelImage At(int row, int column, int offsetX = 0, int offsetY = 0)
    { _definition.SetAnchor(ExcelRangeReference.Cell(row, column), offsetX, offsetY); return this; }
    /// <summary>Sets the image width in pixels.</summary>
    public ExcelImage Width(int value) { _definition.SetWidth(value); return this; }
    /// <summary>Sets the image height in pixels.</summary>
    public ExcelImage Height(int value) { _definition.SetHeight(value); return this; }
    /// <summary>Sets explicit image dimensions in pixels. This takes precedence over aspect-ratio preservation.</summary>
    public ExcelImage Size(int width, int height) { _definition.SetSize(width, height); return this; }
    /// <summary>Preserves the original image aspect ratio when exactly one dimension is configured.</summary>
    public ExcelImage KeepAspectRatio() { _definition.KeepAspectRatio = true; return this; }
    /// <summary>Sets the image's accessible drawing name.</summary>
    public ExcelImage Name(string value) { _definition.SetName(value); return this; }
    /// <summary>Sets alternative text exposed by compatible spreadsheet applications.</summary>
    public ExcelImage AltText(string value) { _definition.SetAltText(value); return this; }
}
