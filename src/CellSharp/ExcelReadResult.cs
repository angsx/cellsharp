namespace CellSharp;

/// <summary>Contains rows successfully read from a worksheet and any data issues found.</summary>
public sealed class ExcelReadResult<T>
{
    internal ExcelReadResult(IEnumerable<T> rows, IEnumerable<ExcelReadError> errors)
    {
        Rows = Array.AsReadOnly(rows.ToArray());
        Errors = Array.AsReadOnly(errors.ToArray());
    }

    /// <summary>Gets rows whose mapped values were all converted successfully.</summary>
    public IReadOnlyList<T> Rows { get; }

    /// <summary>Gets data issues found while reading the worksheet.</summary>
    public IReadOnlyList<ExcelReadError> Errors { get; }

    /// <summary>Gets whether no data issues were found.</summary>
    public bool IsValid => Errors.Count == 0;
}
