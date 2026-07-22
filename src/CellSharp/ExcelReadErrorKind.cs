namespace CellSharp;

/// <summary>Groups XLSX read errors by their programmatic cause.</summary>
public enum ExcelReadErrorKind
{
    /// <summary>A required worksheet header was not found.</summary>
    MissingHeader,

    /// <summary>A non-empty worksheet header was repeated.</summary>
    DuplicateHeader,

    /// <summary>A required cell value was empty.</summary>
    RequiredValue,

    /// <summary>A cell value or converter result could not be converted.</summary>
    Conversion,

    /// <summary>A converted value did not satisfy declarative or custom validation.</summary>
    Validation,
}
