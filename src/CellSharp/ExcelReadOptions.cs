using System.Globalization;

namespace CellSharp;

/// <summary>Configures one XLSX read operation.</summary>
public sealed class ExcelReadOptions
{
    /// <summary>The default maximum compressed XLSX package size: 64 MiB.</summary>
    public const long DefaultMaxPackageBytes = 64L * 1024L * 1024L;

    /// <summary>The default maximum number of XML characters in any one Open XML part: 32 MiB.</summary>
    public const long DefaultMaxCharactersInPart = 32L * 1024L * 1024L;

    /// <summary>The default maximum number of physical rows inspected in one worksheet.</summary>
    public const int DefaultMaxRows = 100000;

    /// <summary>The default maximum number of data errors collected in one worksheet.</summary>
    public const int DefaultMaxErrors = 1000;

    /// <summary>Initializes read options with bounded defaults suitable for untrusted XLSX input.</summary>
    public ExcelReadOptions(
        ExcelInvalidRowPolicy invalidRowPolicy = ExcelInvalidRowPolicy.Skip,
        int maxRows = DefaultMaxRows,
        int maxErrors = DefaultMaxErrors,
        long maxPackageBytes = DefaultMaxPackageBytes,
        long maxCharactersInPart = DefaultMaxCharactersInPart,
        CultureInfo? culture = null,
        bool emptyStringAsNull = false)
    {
        if (!Enum.IsDefined(typeof(ExcelInvalidRowPolicy), invalidRowPolicy))
        {
            throw new ArgumentOutOfRangeException(nameof(invalidRowPolicy));
        }

        if (maxRows < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(maxRows), "The row limit must be greater than zero.");
        }

        if (maxErrors < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(maxErrors), "The error limit must be greater than zero.");
        }

        if (maxPackageBytes < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(maxPackageBytes), "The package-size limit must be greater than zero.");
        }

        if (maxCharactersInPart < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(maxCharactersInPart), "The part-size limit must be greater than zero.");
        }

        InvalidRowPolicy = invalidRowPolicy;
        MaxRows = maxRows;
        MaxErrors = maxErrors;
        MaxPackageBytes = maxPackageBytes;
        MaxCharactersInPart = maxCharactersInPart;
        Culture = culture ?? CultureInfo.InvariantCulture;
        EmptyStringAsNull = emptyStringAsNull;
    }

    /// <summary>Gets the materialization policy for rows containing conversion or validation errors.</summary>
    public ExcelInvalidRowPolicy InvalidRowPolicy { get; }

    /// <summary>Gets the maximum number of physical worksheet rows inspected before the input is rejected.</summary>
    public int MaxRows { get; }

    /// <summary>Gets the maximum number of data errors collected before the input is rejected.</summary>
    public int MaxErrors { get; }

    /// <summary>Gets the maximum compressed XLSX package size accepted from a path or seekable stream.</summary>
    public long MaxPackageBytes { get; }

    /// <summary>Gets the maximum number of XML characters accepted in any one Open XML part.</summary>
    public long MaxCharactersInPart { get; }

    /// <summary>Gets the culture used as a fallback for text-form numeric and date values.</summary>
    /// <remarks>
    /// Defaults to <see cref="CultureInfo.InvariantCulture"/>. Invariant numeric and date conversion is always attempted first;
    /// string, Boolean, and GUID conversion is not affected.
    /// </remarks>
    public CultureInfo Culture { get; }

    /// <summary>Gets whether explicit zero-length text cells are treated as null before scalar conversion.</summary>
    /// <remarks>Defaults to <see langword="false"/> and does not treat whitespace as empty.</remarks>
    public bool EmptyStringAsNull { get; }

    internal static ExcelReadOptions Default { get; } = new();
}
