using System.Buffers.Binary;

namespace CellSharp.Internal;

internal sealed class WorksheetImageDefinition
{
    internal WorksheetImageDefinition(byte[] data, ExcelImageFormat format, int width, int height)
    { Data = data; Format = format; OriginalWidth = width; OriginalHeight = height; }
    internal byte[] Data { get; }
    internal ExcelImageFormat Format { get; }
    internal int OriginalWidth { get; }
    internal int OriginalHeight { get; }
    internal ExcelRangeReference? Anchor { get; private set; }
    internal int OffsetX { get; private set; }
    internal int OffsetY { get; private set; }
    internal int? RequestedWidth { get; private set; }
    internal int? RequestedHeight { get; private set; }
    internal bool ExplicitSize { get; private set; }
    internal bool KeepAspectRatio { get; set; }
    internal string? Name { get; private set; }
    internal string? AltText { get; private set; }
    internal void SetAnchor(ExcelRangeReference anchor, int offsetX, int offsetY)
    { if (offsetX < 0) ThrowInvalidOffset(nameof(offsetX)); if (offsetY < 0) ThrowInvalidOffset(nameof(offsetY)); Anchor = anchor; OffsetX = offsetX; OffsetY = offsetY; }
    internal void SetWidth(int value) { ValidatePixels(value, nameof(value)); RequestedWidth = value; ExplicitSize = false; }
    internal void SetHeight(int value) { ValidatePixels(value, nameof(value)); RequestedHeight = value; ExplicitSize = false; }
    internal void SetSize(int width, int height) { ValidatePixels(width, nameof(width)); ValidatePixels(height, nameof(height)); RequestedWidth = width; RequestedHeight = height; ExplicitSize = true; }
    internal void SetName(string value) { if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("An image name is required.", nameof(value)); Name = value; }
    internal void SetAltText(string value) { if (value is null) throw new ArgumentNullException(nameof(value)); AltText = value; }
    internal (int width, int height) ResolveSize()
    {
        if (Anchor is null) throw new InvalidOperationException("An image must be positioned with At(...).");
        if (ExplicitSize) return (RequestedWidth!.Value, RequestedHeight!.Value);
        if (RequestedWidth is not null && RequestedHeight is not null)
        {
            if (KeepAspectRatio) throw new InvalidOperationException("KeepAspectRatio requires exactly one configured dimension; Size(width, height) is already explicit.");
            return (RequestedWidth.Value, RequestedHeight.Value);
        }
        if (RequestedWidth is not null) return (RequestedWidth.Value, KeepAspectRatio ? Scale(RequestedWidth.Value, OriginalWidth, OriginalHeight) : OriginalHeight);
        if (RequestedHeight is not null) return (KeepAspectRatio ? Scale(RequestedHeight.Value, OriginalHeight, OriginalWidth) : OriginalWidth, RequestedHeight.Value);
        return (OriginalWidth, OriginalHeight);
    }
    private static int Scale(int value, int source, int target) => checked((int)Math.Max(1, Math.Round((double)value * target / source, MidpointRounding.AwayFromZero)));
    private static void ValidatePixels(int value, string parameterName) { if (value <= 0) throw new ArgumentOutOfRangeException(parameterName, "Image dimensions must be greater than zero."); }
    private static void ThrowInvalidOffset(string parameterName) => throw new ArgumentOutOfRangeException(parameterName, "Image offsets cannot be negative.");
}

internal static class ExcelImageData
{
    internal static WorksheetImageDefinition FromPath(string path)
    {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("An image file path is required.", nameof(path));
        if (Directory.Exists(path)) throw new ArgumentException("The image path must identify a file, not a directory.", nameof(path));
        var format = Path.GetExtension(path).ToLowerInvariant() switch { ".png" => ExcelImageFormat.Png, ".jpg" or ".jpeg" => ExcelImageFormat.Jpeg, _ => throw new NotSupportedException("Only PNG and JPEG images are supported.") };
        var bytes = File.ReadAllBytes(path);
        return FromBytes(bytes, format, "image file");
    }
    internal static WorksheetImageDefinition FromStream(Stream stream, ExcelImageFormat format)
    {
        if (stream is null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("The image stream must be readable.", nameof(stream));
        using var copy = new MemoryStream();
        stream.CopyTo(copy);
        return FromBytes(copy.ToArray(), format, "image stream");
    }
    private static WorksheetImageDefinition FromBytes(byte[] bytes, ExcelImageFormat format, string source)
    {
        if (bytes.Length == 0) throw new ArgumentException($"The {source} is empty.");
        var dimensions = format switch
        {
            ExcelImageFormat.Png => PngDimensions(bytes, source),
            ExcelImageFormat.Jpeg => JpegDimensions(bytes, source),
            _ => throw new ArgumentOutOfRangeException(nameof(format)),
        };
        return new WorksheetImageDefinition(bytes, format, dimensions.width, dimensions.height);
    }
    private static (int width, int height) PngDimensions(byte[] bytes, string source)
    {
        if (bytes.Length < 24 || !bytes.AsSpan(0, 8).SequenceEqual(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 })) throw new ArgumentException($"The {source} is not a valid PNG image.");
        var width = BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(16, 4)); var height = BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(20, 4));
        if (width <= 0 || height <= 0) throw new ArgumentException("The PNG image has invalid dimensions.");
        return (width, height);
    }
    private static (int width, int height) JpegDimensions(byte[] bytes, string source)
    {
        if (bytes.Length < 4 || bytes[0] != 0xFF || bytes[1] != 0xD8) throw new ArgumentException($"The {source} is not a valid JPEG image.");
        for (var index = 2; index + 8 < bytes.Length;)
        {
            if (bytes[index] != 0xFF) { index++; continue; }
            while (index < bytes.Length && bytes[index] == 0xFF) index++;
            if (index >= bytes.Length) break;
            var marker = bytes[index++];
            if (marker is 0xD8 or 0xD9 || marker == 0x01 || marker is >= 0xD0 and <= 0xD7) continue;
            if (index + 1 >= bytes.Length) break;
            var length = (bytes[index] << 8) | bytes[index + 1];
            if (length < 2 || index + length > bytes.Length) break;
            if (marker is >= 0xC0 and <= 0xC3 or >= 0xC5 and <= 0xC7 or >= 0xC9 and <= 0xCB or >= 0xCD and <= 0xCF)
            {
                var height = (bytes[index + 3] << 8) | bytes[index + 4]; var width = (bytes[index + 5] << 8) | bytes[index + 6];
                if (width <= 0 || height <= 0) break;
                return (width, height);
            }
            index += length;
        }
        throw new ArgumentException($"The {source} does not contain supported JPEG dimensions.");
    }
}
