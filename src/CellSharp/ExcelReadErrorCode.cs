namespace CellSharp;

/// <summary>Identifies a data issue found while reading an XLSX worksheet.</summary>
public enum ExcelReadErrorCode
{
    /// <summary>A required worksheet header was not found.</summary>
    MissingHeader,
    /// <summary>More than one worksheet header maps to the same name.</summary>
    DuplicateHeader,
    /// <summary>A required non-nullable cell value was blank.</summary>
    RequiredValueMissing,
    /// <summary>A cell value could not be converted to its target CLR type.</summary>
    InvalidValue,
    /// <summary>A converted value did not satisfy a configured validation rule.</summary>
    ValidationFailed,
}
