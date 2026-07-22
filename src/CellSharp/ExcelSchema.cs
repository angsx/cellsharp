using System.Collections.ObjectModel;
using System.Linq.Expressions;
using System.Reflection;
using CellSharp.Internal;

namespace CellSharp;

/// <summary>Describes the XLSX representation of a .NET type.</summary>
public sealed class ExcelSchema<T>
{
    internal ExcelSchema(
        IEnumerable<SchemaColumn> columns,
        string sheetName,
        SchemaTableDefinition? table,
        WorksheetSettingsDefinition worksheetSettings)
    {
        Columns = new ReadOnlyCollection<SchemaColumn>(columns.ToArray());
        SheetName = sheetName;
        Table = table;
        WorksheetSettings = worksheetSettings;
    }

    internal IReadOnlyList<SchemaColumn> Columns { get; }

    internal SchemaTableDefinition? Table { get; }

    internal WorksheetSettingsDefinition WorksheetSettings { get; }

    /// <summary>Gets the worksheet name used when this schema reads, writes, or creates a template.</summary>
    public string SheetName { get; }
}

/// <summary>Builds an immutable typed XLSX schema.</summary>
public sealed class ExcelSchemaBuilder<T>
{
    private readonly List<SchemaColumn> _columns = new();
    private readonly HashSet<PropertyInfo> _attributeColumns = new();
    private string _sheetName = ExcelSheetName.Default;
    private SchemaTableDefinition? _table;
    private bool _autoFilter;
    private int? _freezeRows;
    private int? _freezeColumns;
    private bool? _landscape;
    private int? _fitToWidth;
    private int? _fitToHeight;
    private bool _repeatHeaderRowOnPrint;
    private bool _printGridlines;

    /// <summary>Sets the worksheet name used when the built schema reads, writes, or creates a template.</summary>
    public ExcelSchemaBuilder<T> SheetName(string sheetName)
    {
        _sheetName = ExcelSheetName.Validate(sheetName, nameof(sheetName));
        return this;
    }

    /// <summary>Exports this schema as a native Excel Table with optional name and built-in style.</summary>
    public ExcelSchemaBuilder<T> AsTable(string? name = null, string? style = "TableStyleMedium2")
    {
        if (_table is not null)
        {
            throw new InvalidOperationException("A table is already configured for this schema.");
        }

        _table = new SchemaTableDefinition(name, style);
        return this;
    }

    /// <summary>Adds a native worksheet AutoFilter unless the schema is exported as an Excel Table.</summary>
    public ExcelSchemaBuilder<T> AutoFilter()
    {
        _autoFilter = true;
        return this;
    }

    /// <summary>Freezes the specified number of top rows and left columns. A zero value leaves that axis unfrozen.</summary>
    public ExcelSchemaBuilder<T> FreezePanes(int rows, int columns)
    {
#if NET8_0_OR_GREATER
        ArgumentOutOfRangeException.ThrowIfNegative(rows);
        ArgumentOutOfRangeException.ThrowIfNegative(columns);
#else
        if (rows < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(rows));
        }

        if (columns < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(columns));
        }
#endif

        _freezeRows = rows;
        _freezeColumns = columns;
        return this;
    }

    /// <summary>Uses landscape orientation when printing this worksheet.</summary>
    public ExcelSchemaBuilder<T> Landscape()
    {
        _landscape = true;
        return this;
    }

    /// <summary>Uses portrait orientation when printing this worksheet.</summary>
    public ExcelSchemaBuilder<T> Portrait()
    {
        _landscape = false;
        return this;
    }

    /// <summary>Fits printed pages to the supplied width and height. Zero means automatic for that dimension.</summary>
    public ExcelSchemaBuilder<T> FitToPage(int width, int height)
    {
#if NET8_0_OR_GREATER
        ArgumentOutOfRangeException.ThrowIfNegative(width);
        ArgumentOutOfRangeException.ThrowIfNegative(height);
#else
        if (width < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(width));
        }

        if (height < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(height));
        }
#endif

        _fitToWidth = width;
        _fitToHeight = height;
        return this;
    }

    /// <summary>Repeats worksheet row one as the header row on every printed page.</summary>
    public ExcelSchemaBuilder<T> RepeatHeaderRowOnPrint()
    {
        _repeatHeaderRowOnPrint = true;
        return this;
    }

    /// <summary>Prints worksheet gridlines without changing their on-screen visibility.</summary>
    public ExcelSchemaBuilder<T> PrintGridlines()
    {
        _printGridlines = true;
        return this;
    }

    /// <summary>Adds a property to the schema.</summary>
    public ExcelSchemaBuilder<T> Column<TValue>(
        Expression<Func<T, TValue>> property,
        Action<ExcelColumnBuilder<T, TValue>>? configure = null)
    {
        var propertyInfo = PropertyFrom(property);
        ValidateSchemaProperty(propertyInfo, nameof(property));

        var existingIndex = _columns.FindIndex(column => column.Property == propertyInfo);
        if (existingIndex >= 0 && !_attributeColumns.Remove(propertyInfo))
        {
            throw new ArgumentException($"Property '{propertyInfo.Name}' is already configured.", nameof(property));
        }

        var existing = existingIndex >= 0 ? _columns[existingIndex] : null;
        if (existingIndex >= 0)
        {
            _columns.RemoveAt(existingIndex);
        }

        var builder = existing is null
            ? new ExcelColumnBuilder<T, TValue>(propertyInfo)
            : new ExcelColumnBuilder<T, TValue>(propertyInfo, existing);
        configure?.Invoke(builder);
        var column = builder.ToColumn();

        EnsureUniqueHeader(column, nameof(property));
        if (existingIndex >= 0)
        {
            _columns.Insert(existingIndex, column);
        }
        else
        {
            _columns.Add(column);
        }

        return this;
    }

    internal void AddAttributeColumn(PropertyInfo property, ExcelColumnAttribute? attribute, bool ignore)
    {
        ValidateSchemaProperty(property, nameof(property));

        var builder = new ExcelColumnBuilder<T, object?>(property);
        if (attribute?.Header is not null)
        {
            builder.Header(attribute.Header);
        }

        if (attribute?.Optional == true)
        {
            builder.Optional();
        }

        if (attribute?.Format is not null)
        {
            builder.Format(attribute.Format);
        }

        if (ignore)
        {
            builder.Ignore();
        }

        var column = builder.ToColumn(validateSupportedType: false);
        EnsureUniqueHeader(column, nameof(property));
        _columns.Add(column);
        _attributeColumns.Add(property);
    }

    /// <summary>Creates an immutable schema that can be reused for read and write operations.</summary>
    public ExcelSchema<T> Build()
    {
        if (_columns.All(column => column.IsIgnored))
        {
            throw new InvalidOperationException("A schema must include at least one column.");
        }

        ValidateSupportedColumnTypes();
        ValidateRuntimeMappings();
        return new ExcelSchema<T>(
            _columns,
            _sheetName,
            _table,
            new WorksheetSettingsDefinition(
                _autoFilter,
                _freezeRows,
                _freezeColumns,
                _landscape,
                _fitToWidth,
                _fitToHeight,
                _repeatHeaderRowOnPrint,
                _printGridlines));
    }

    private void ValidateRuntimeMappings()
    {
        var columns = _columns.Where(column => !column.IsIgnored).ToArray();
        var duplicateColumnNumber = columns
            .Where(column => column.SourceColumnNumber is not null)
            .GroupBy(column => column.SourceColumnNumber!.Value)
            .FirstOrDefault(group => group.Count() > 1);
        if (duplicateColumnNumber is not null)
        {
            throw new ArgumentException(
                $"Runtime column number '{duplicateColumnNumber.Key}' is mapped to more than one property.");
        }

        var duplicateSourceHeader = columns
            .Where(column => column.SourceHeader is not null)
            .GroupBy(column => column.SourceHeader!, StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault(group => group.Count() > 1);
        if (duplicateSourceHeader is not null)
        {
            throw new ArgumentException(
                $"Runtime header '{duplicateSourceHeader.Key}' is mapped to more than one property.");
        }

        foreach (var column in columns.Where(column => column.SourceHeader is not null))
        {
            var conflictingHeader = columns.FirstOrDefault(other => other.Property != column.Property
                && string.Equals(other.Header, column.SourceHeader, StringComparison.OrdinalIgnoreCase));
            if (conflictingHeader is not null)
            {
                throw new ArgumentException(
                    $"Runtime header '{column.SourceHeader}' for property '{column.Property.Name}' conflicts with the schema header for property '{conflictingHeader.Property.Name}'.");
            }
        }
    }

    private void ValidateSupportedColumnTypes()
    {
        var unsupported = _columns.FirstOrDefault(column =>
            !column.IsIgnored && !CellValueConverter.Supports(column.Property.PropertyType) && column.Converter is null);
        if (unsupported is not null)
        {
            throw new NotSupportedException(
                $"Property '{unsupported.Property.Name}' on type '{typeof(T).FullName}' has unsupported type '{unsupported.Property.PropertyType.FullName}'. Configure ConvertWith to use it in a schema.");
        }
    }

    private static PropertyInfo PropertyFrom<TValue>(Expression<Func<T, TValue>> expression)
    {
        if (expression is null)
        {
#if NET8_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(expression);
#else
            throw new ArgumentNullException(nameof(expression));
#endif
        }

        if (expression.Body is not MemberExpression { Member: PropertyInfo property, Expression: ParameterExpression })
        {
            throw new ArgumentException("A schema column must select a direct property access.", nameof(expression));
        }

        return property;
    }

    private static void ValidateSchemaProperty(PropertyInfo propertyInfo, string parameterName)
    {
        if (!propertyInfo.CanRead || propertyInfo.SetMethod?.IsPublic != true)
        {
            throw new ArgumentException(
                $"Property '{propertyInfo.Name}' must be publicly readable and writable to be used in a schema.",
                parameterName);
        }
    }

    private void EnsureUniqueHeader(SchemaColumn column, string parameterName)
    {
        if (!column.IsIgnored && _columns.Any(existing =>
                !existing.IsIgnored && string.Equals(existing.Header, column.Header, StringComparison.OrdinalIgnoreCase)))
        {
            throw new ArgumentException($"Header '{column.Header}' is already configured.", parameterName);
        }
    }
}

/// <summary>Configures one schema column while it is being built.</summary>
public sealed class ExcelColumnBuilder<T, TValue>
{
    private readonly PropertyInfo _property;
    private string _header;
    private bool _required = true;
    private bool _ignored;
    private int? _sourceColumnNumber;
    private string? _sourceHeader;
    private string? _format;
    private double? _width;
    private ExcelHorizontalAlignment? _alignment;
    private readonly List<ValidationRule> _validations = new();
    private DeclarativeValidationRule? _declarativeValidation;
    private ValueConverterDefinition? _converter;
    private Func<ExcelFormulaContext, string>? _formula;

    internal ExcelColumnBuilder(PropertyInfo property)
    {
        _property = property;
        _header = property.Name;
    }

    internal ExcelColumnBuilder(PropertyInfo property, SchemaColumn column)
    {
        _property = property;
        _header = column.Header;
        _required = column.IsRequired;
        _ignored = column.IsIgnored;
        _sourceColumnNumber = column.SourceColumnNumber;
        _sourceHeader = column.SourceHeader;
        _format = column.Format;
        _width = column.Width;
        _alignment = column.Alignment;
        _validations.AddRange(column.Validations);
        _declarativeValidation = column.DeclarativeValidation;
        _converter = column.Converter;
        _formula = column.Formula;
    }

    /// <summary>Sets the worksheet header for this property.</summary>
    public ExcelColumnBuilder<T, TValue> Header(string header)
    {
        if (string.IsNullOrWhiteSpace(header))
        {
            throw new ArgumentException("A header is required.", nameof(header));
        }

        _header = header;
        return this;
    }

    /// <summary>Allows this header to be absent when reading.</summary>
    public ExcelColumnBuilder<T, TValue> Optional()
    {
        _required = false;
        return this;
    }

    /// <summary>Excludes this property from both reading and writing.</summary>
    public ExcelColumnBuilder<T, TValue> Ignore()
    {
        _ignored = true;
        return this;
    }

    /// <summary>Maps this property from a one-based worksheet column number when reading.</summary>
    public ExcelColumnBuilder<T, TValue> MapFromColumn(int columnNumber)
    {
        if (columnNumber < 1 || columnNumber > 16384)
        {
            throw new ArgumentOutOfRangeException(nameof(columnNumber), "Column numbers must be between 1 and Excel's maximum column 16384 (XFD).");
        }

        if (_sourceColumnNumber is not null)
        {
            throw new InvalidOperationException("A runtime column number is already configured for this property.");
        }

        _sourceColumnNumber = columnNumber;
        return this;
    }

    /// <summary>Maps this property from the supplied worksheet header when reading.</summary>
    public ExcelColumnBuilder<T, TValue> MapFromHeader(string header)
    {
        if (string.IsNullOrWhiteSpace(header))
        {
            throw new ArgumentException("A runtime header is required.", nameof(header));
        }

        if (_sourceHeader is not null)
        {
            throw new InvalidOperationException("A runtime header is already configured for this property.");
        }

        _sourceHeader = header;
        return this;
    }

    /// <summary>Sets the Excel number or date format code used when exporting this column.</summary>
    public ExcelColumnBuilder<T, TValue> Format(string format)
    {
        if (string.IsNullOrWhiteSpace(format))
        {
            throw new ArgumentException("An Excel format code is required.", nameof(format));
        }

        _format = format;
        return this;
    }

    /// <summary>Sets the explicit Excel column width. Explicit widths take precedence over AutoFitColumns.</summary>
    public ExcelColumnBuilder<T, TValue> Width(double width)
    {
        if (width <= 0D || width > 255D)
        {
            throw new ArgumentOutOfRangeException(nameof(width), "Width must be greater than zero and no greater than 255.");
        }

        _width = width;
        return this;
    }

    /// <summary>Sets the horizontal alignment used for exported data cells in this column.</summary>
    public ExcelColumnBuilder<T, TValue> Align(ExcelHorizontalAlignment alignment)
    {
        if (!Enum.IsDefined(typeof(ExcelHorizontalAlignment), alignment))
        {
            throw new ArgumentOutOfRangeException(nameof(alignment));
        }

        _alignment = alignment;
        return this;
    }

    /// <summary>Uses a reusable bidirectional converter for this column.</summary>
    public ExcelColumnBuilder<T, TValue> ConvertWith<TCellValue>(IExcelValueConverter<TValue, TCellValue> converter)
    {
        if (converter is null)
        {
#if NET8_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(converter);
#else
            throw new ArgumentNullException(nameof(converter));
#endif
        }

        if (!CellValueConverter.Supports(typeof(TCellValue)))
        {
            throw new NotSupportedException(
                $"Converter cell type '{typeof(TCellValue).FullName}' is not supported by CellSharp.");
        }

        if (_converter is not null)
        {
            throw new InvalidOperationException("Only one converter can be configured for a column.");
        }

        _converter = new ValueConverterDefinition<TValue, TCellValue>(converter);
        return this;
    }

    /// <summary>Writes a native Excel formula for this column on each exported data row.</summary>
    /// <remarks>The returned formula is executable spreadsheet content and must not be built from untrusted input.</remarks>
    public ExcelColumnBuilder<T, TValue> Formula(Func<ExcelFormulaContext, string> formula)
    {
        if (formula is null)
        {
#if NET8_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(formula);
#else
            throw new ArgumentNullException(nameof(formula));
#endif
        }

        if (_formula is not null)
        {
            throw new InvalidOperationException("Only one formula can be configured for a column.");
        }

        _formula = formula;
        return this;
    }

    /// <summary>Restricts a string column to the supplied non-empty set of values.</summary>
    public ExcelColumnBuilder<T, TValue> AllowedValues(params string[] values)
    {
        if (UnderlyingType() != typeof(string))
        {
            throw new InvalidOperationException("AllowedValues can only be configured for string columns.");
        }

        if (values is null || values.Length == 0)
        {
            throw new ArgumentException("At least one allowed value is required.", nameof(values));
        }

        if (values.Any(string.IsNullOrEmpty))
        {
            throw new ArgumentException("Allowed values cannot be null or empty.", nameof(values));
        }

        if (values.Length > 1048576)
        {
            throw new ArgumentOutOfRangeException(nameof(values), "Allowed values cannot exceed Excel's maximum of 1048576 rows.");
        }

        SetDeclarativeValidation(new AllowedValuesValidationRule(values.Distinct(StringComparer.Ordinal)));
        return this;
    }

    /// <summary>Restricts a numeric column to an inclusive range.</summary>
    public ExcelColumnBuilder<T, TValue> Range(TValue minimum, TValue maximum)
    {
        if (!IsNumeric(UnderlyingType()))
        {
            throw new InvalidOperationException("Range can only be configured for numeric columns.");
        }

        if (minimum is null || maximum is null)
        {
            throw new ArgumentException("Range bounds cannot be null.");
        }

        var min = (object)minimum;
        var max = (object)maximum;
        if (((IComparable)min).CompareTo(max) > 0)
        {
            throw new ArgumentException("The minimum range value cannot be greater than the maximum range value.");
        }

        SetDeclarativeValidation(new NumericRangeValidationRule(min, max));
        return this;
    }

    /// <summary>Restricts a DateTime column to an inclusive date range.</summary>
    public ExcelColumnBuilder<T, TValue> DateBetween(DateTime minimum, DateTime maximum)
    {
        if (UnderlyingType() != typeof(DateTime))
        {
            throw new InvalidOperationException("DateBetween can only be configured for DateTime columns.");
        }

        if (minimum > maximum)
        {
            throw new ArgumentException("The minimum date cannot be greater than the maximum date.");
        }

        SetDeclarativeValidation(new DateRangeValidationRule(minimum, maximum));
        return this;
    }

    /// <summary>Adds a validation rule evaluated after the cell value has been converted.</summary>
    public ExcelColumnBuilder<T, TValue> Validate(Func<TValue, bool> predicate, string message)
    {
        if (predicate is null)
        {
#if NET8_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(predicate);
#else
            throw new ArgumentNullException(nameof(predicate));
#endif
        }

        if (string.IsNullOrWhiteSpace(message))
        {
            throw new ArgumentException("A validation message is required.", nameof(message));
        }

        _validations.Add(new ValidationRule(value => predicate((TValue)value!), message));
        return this;
    }

    internal SchemaColumn ToColumn(bool validateSupportedType = true)
    {
        if (validateSupportedType && !CellValueConverter.Supports(_property.PropertyType) && _converter is null)
        {
            throw new NotSupportedException(
                $"Property '{_property.Name}' on type '{typeof(T).FullName}' has unsupported type '{_property.PropertyType.FullName}'. Configure ConvertWith to use it in a schema.");
        }

        if (_ignored && (_declarativeValidation is not null || _formula is not null))
        {
            throw new InvalidOperationException("Ignored columns cannot have declarative validation or formulas.");
        }

        if (_formula is not null && _declarativeValidation is not null)
        {
            throw new InvalidOperationException("Formula columns cannot have declarative validation.");
        }

        if (_ignored && (_sourceColumnNumber is not null || _sourceHeader is not null))
        {
            throw new InvalidOperationException("Ignored columns cannot have runtime mappings.");
        }

        return new SchemaColumn(
            _property,
            _header,
            _required,
            _ignored,
            _validations,
            _declarativeValidation,
            _converter,
            _format,
            _width,
            _alignment,
            _sourceColumnNumber,
            _sourceHeader,
            _formula);
    }

    private void SetDeclarativeValidation(DeclarativeValidationRule validation)
    {
        if (_declarativeValidation is not null)
        {
            throw new InvalidOperationException("Only one declarative validation can be configured for a column.");
        }

        _declarativeValidation = validation;
    }

    private Type UnderlyingType() => Nullable.GetUnderlyingType(_property.PropertyType) ?? _property.PropertyType;

    private static bool IsNumeric(Type type) => type == typeof(byte)
        || type == typeof(sbyte)
        || type == typeof(short)
        || type == typeof(ushort)
        || type == typeof(int)
        || type == typeof(uint)
        || type == typeof(long)
        || type == typeof(ulong)
        || type == typeof(float)
        || type == typeof(double)
        || type == typeof(decimal);
}
