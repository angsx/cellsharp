namespace CellSharp;

/// <summary>Describes one worksheet data issue without exposing Open XML details.</summary>
public sealed class ExcelReadError
{
    internal ExcelReadError(
        string sheetName,
        uint rowNumber,
        int? columnNumber,
        string? cellReference,
        string? header,
        string? propertyName,
        string? value,
        Type? expectedType,
        ExcelReadErrorCode code,
        string message)
    {
        SheetName = sheetName;
        RowNumber = rowNumber;
        ColumnNumber = columnNumber;
        CellReference = cellReference;
        Header = header;
        PropertyName = propertyName;
        Value = value;
        ExpectedType = expectedType;
        Code = code;
        Message = message;
    }

    /// <summary>Gets the worksheet name where the issue was found.</summary>
    public string SheetName { get; }

    /// <summary>Gets the one-based visible Excel row number.</summary>
    public uint RowNumber { get; }

    /// <summary>Gets the one-based visible Excel row number.</summary>
    public uint Row => RowNumber;

    /// <summary>Gets the one-based physical Excel column number, when applicable.</summary>
    public int? ColumnNumber { get; }

    /// <summary>Gets the one-based physical Excel column number, when applicable.</summary>
    public int? Column => ColumnNumber;

    /// <summary>Gets the A1-style cell reference, when applicable.</summary>
    public string? CellReference { get; }

    /// <summary>Gets the worksheet header associated with the issue, when applicable.</summary>
    public string? Header { get; }

    /// <summary>Gets the CLR property targeted by this error, when applicable.</summary>
    public string? PropertyName { get; }

    /// <summary>Gets the untrusted raw textual cell value used by the import pipeline, when applicable.</summary>
    public string? Value { get; }

    /// <summary>Gets the untrusted original cell value before conversion, when applicable.</summary>
    public object? RawValue => Value;

    /// <summary>Gets the expected CLR type, when conversion was attempted.</summary>
    public Type? ExpectedType { get; }

    /// <summary>Gets the detailed compatible error code.</summary>
    public ExcelReadErrorCode Code { get; }

    /// <summary>Gets the programmatic category of this error.</summary>
    public ExcelReadErrorKind Kind => Code switch
    {
        ExcelReadErrorCode.MissingHeader => ExcelReadErrorKind.MissingHeader,
        ExcelReadErrorCode.DuplicateHeader => ExcelReadErrorKind.DuplicateHeader,
        ExcelReadErrorCode.RequiredValueMissing => ExcelReadErrorKind.RequiredValue,
        ExcelReadErrorCode.InvalidValue => ExcelReadErrorKind.Conversion,
        ExcelReadErrorCode.ValidationFailed => ExcelReadErrorKind.Validation,
        _ => throw new InvalidOperationException($"Unknown read error code '{Code}'."),
    };

    /// <summary>Gets a human-readable explanation of the issue.</summary>
    public string Message { get; }
}
