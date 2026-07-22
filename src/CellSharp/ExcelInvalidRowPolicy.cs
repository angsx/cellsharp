namespace CellSharp;

/// <summary>Controls whether rows with cell conversion or validation errors are returned.</summary>
public enum ExcelInvalidRowPolicy
{
    /// <summary>Excludes rows containing one or more data errors.</summary>
    Skip,

    /// <summary>Includes rows containing data errors with successfully converted values and CLR defaults for failed conversions.</summary>
    Include,
}
