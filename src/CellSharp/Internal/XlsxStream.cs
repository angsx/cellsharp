using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace CellSharp.Internal;

internal static class XlsxStream
{
    internal static SpreadsheetDocument Create(Stream stream)
    {
        Validate(stream, requireWrite: true);
        stream.Position = 0;
        stream.SetLength(0);
        return SpreadsheetDocument.Create(new LeaveOpenStream(stream), SpreadsheetDocumentType.Workbook);
    }

    internal static SpreadsheetDocument Open(Stream stream, ExcelReadOptions? options = null)
    {
        Validate(stream, requireWrite: false);
        var readOptions = options ?? ExcelReadOptions.Default;
        ValidatePackageLength(stream.Length, readOptions.MaxPackageBytes);
        stream.Position = 0;
        return SpreadsheetDocument.Open(new LeaveOpenStream(stream), false, OpenSettings(readOptions));
    }

    internal static SpreadsheetDocument Open(string path, ExcelReadOptions? options = null)
    {
        var readOptions = options ?? ExcelReadOptions.Default;
        ValidatePackageLength(new FileInfo(path).Length, readOptions.MaxPackageBytes);
        return SpreadsheetDocument.Open(path, false, OpenSettings(readOptions));
    }

    internal static void CompleteWrite(Stream stream) => stream.Position = stream.Length;

    internal static void ValidateReadable(Stream stream) => Validate(stream, requireWrite: false);

    internal static void ValidateWritable(Stream stream) => Validate(stream, requireWrite: true);

    private static OpenSettings OpenSettings(ExcelReadOptions options) => new()
    {
        AutoSave = false,
        MaxCharactersInPart = options.MaxCharactersInPart,
    };

    private static void ValidatePackageLength(long length, long maximum)
    {
        if (length > maximum)
        {
            throw new InvalidDataException($"The XLSX package exceeds the configured limit of {maximum} bytes.");
        }
    }

    private static void Validate(Stream stream, bool requireWrite)
    {
        if (stream is null)
        {
#if NET8_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(stream);
#else
            throw new ArgumentNullException(nameof(stream));
#endif
        }

        if (requireWrite ? !stream.CanWrite : !stream.CanRead)
        {
            throw new ArgumentException(
                requireWrite ? "The stream must be writable." : "The stream must be readable.",
                nameof(stream));
        }

        if (!stream.CanSeek)
        {
            throw new NotSupportedException("The stream must support seeking.");
        }
    }
}

internal sealed class LeaveOpenStream : Stream
{
    private readonly Stream _inner;

    internal LeaveOpenStream(Stream inner) => _inner = inner;

    public override bool CanRead => _inner.CanRead;

    public override bool CanSeek => _inner.CanSeek;

    public override bool CanWrite => _inner.CanWrite;

    public override bool CanTimeout => _inner.CanTimeout;

    public override long Length => _inner.Length;

    public override long Position
    {
        get => _inner.Position;
        set => _inner.Position = value;
    }

    public override int ReadTimeout
    {
        get => _inner.ReadTimeout;
        set => _inner.ReadTimeout = value;
    }

    public override int WriteTimeout
    {
        get => _inner.WriteTimeout;
        set => _inner.WriteTimeout = value;
    }

    public override void Flush() => _inner.Flush();

    public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);

    public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);

    public override void SetLength(long value) => _inner.SetLength(value);

    public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);

    protected override void Dispose(bool disposing)
    {
        // The caller owns the underlying stream.
        base.Dispose(disposing);
    }
}
