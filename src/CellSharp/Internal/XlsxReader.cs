using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CellSharp.Internal;

internal sealed class XlsxReader<T>
{
    private const int MaximumExcelColumn = 16384;
    private const uint MaximumExcelRow = 1048576U;

    private static readonly HashSet<uint> BuiltInDateFormats = new()
    {
        14U,
        15U,
        16U,
        17U,
        18U,
        19U,
        20U,
        21U,
        22U,
    };

    private readonly string? _path;
    private readonly Stream? _stream;

    internal XlsxReader(string path)
    {
        _path = Path.GetFullPath(path);
    }

    internal XlsxReader(Stream stream)
    {
        _stream = stream;
    }

    internal ExcelReadResult<T> Read(
        ExcelSchema<T>? schema = null,
        ExcelSchemaOverlay<T>? overlay = null,
        ExcelReadOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var readOptions = options ?? ExcelReadOptions.Default;
        if (_stream is not null)
        {
            using var streamDocument = XlsxStream.Open(_stream, readOptions);
            return ReadDocument(streamDocument, schema, overlay, readOptions, "the supplied stream", cancellationToken);
        }

        if (!File.Exists(_path))
        {
            throw new FileNotFoundException($"The file '{_path}' does not exist.", _path);
        }

        using var document = XlsxStream.Open(_path!, readOptions);
        return ReadDocument(document, schema, overlay, readOptions, _path!, cancellationToken);
    }

    private static ExcelReadResult<T> ReadDocument(
        SpreadsheetDocument document,
        ExcelSchema<T>? schema,
        ExcelSchemaOverlay<T>? overlay,
        ExcelReadOptions options,
        string source,
        CancellationToken cancellationToken)
    {
        var workbookPart = document.WorkbookPart ?? throw new InvalidOperationException("The XLSX file has no workbook part.");
        var workbook = workbookPart.Workbook ?? throw new InvalidOperationException("The XLSX file has no workbook.");
        var sheets = workbook.Sheets?.Elements<Sheet>().ToArray();
        if (sheets is null || sheets.Length == 0)
        {
            throw new InvalidOperationException("The XLSX file has no readable worksheet.");
        }

        var sheet = schema is null
            ? sheets[0]
            : sheets.FirstOrDefault(candidate => string.Equals(
                    candidate.Name?.Value,
                    schema.SheetName,
                    StringComparison.OrdinalIgnoreCase))
                ?? throw new InvalidOperationException(
                    $"Worksheet '{schema.SheetName}' was not found in '{source}'.");
        if (XlsxDataValidationWriter.IsInternalSheet(sheet))
        {
            throw new InvalidOperationException("CellSharp internal worksheets cannot be read as application data.");
        }

        return Read(workbookPart, sheet, schema, overlay, options, cancellationToken);
    }

    internal static ExcelReadResult<T> Read(
        WorkbookPart workbookPart,
        Sheet sheet,
        ExcelSchema<T>? schema,
        ExcelSchemaOverlay<T>? overlay = null,
        ExcelReadOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var readOptions = options ?? ExcelReadOptions.Default;
        var metadata = ImportType<T>.CreateMetadata(schema, overlay);
        var relationshipId = sheet.Id?.Value
            ?? throw new InvalidOperationException("The XLSX worksheet has no relationship identifier.");
        var worksheetPart = workbookPart.GetPartById(relationshipId) as WorksheetPart
            ?? throw new InvalidOperationException("The XLSX file has no readable worksheet.");
        var sheetName = sheet.Name?.Value ?? "Sheet1";
        var worksheet = worksheetPart.Worksheet
            ?? throw new InvalidOperationException("The XLSX worksheet part has no worksheet.");
        var rows = worksheet.GetFirstChild<SheetData>()?.Elements<Row>() ?? Enumerable.Empty<Row>();
        var sharedStrings = SharedStrings(workbookPart.SharedStringTablePart, cancellationToken);
        var errors = new List<ExcelReadError>();
        var successfulRows = new List<T>();
        var scan = ScanRows(rows, sharedStrings, readOptions.MaxRows, cancellationToken, schema is null ? null : metadata.Properties);
        var header = scan.Header;
        var usedColumns = scan.UsedColumns;

        var mapping = CreateHeaderMapping(
            header,
            metadata.Properties,
            usedColumns,
            sharedStrings,
            sheetName,
            errors,
            readOptions.MaxErrors);

        if (errors.Count > 0)
        {
            return new ExcelReadResult<T>(successfulRows, errors);
        }

        if (header is null)
        {
            return new ExcelReadResult<T>(successfulRows, errors);
        }

        var dateStyles = DateStyles(workbookPart.WorkbookStylesPart, cancellationToken);
        var fallbackRowNumber = header.RowIndex?.Value ?? 1U;
        foreach (var row in RowsAfter(rows, header))
        {
            cancellationToken.ThrowIfCancellationRequested();
            fallbackRowNumber++;
            var rowNumber = row.RowIndex?.Value ?? fallbackRowNumber;
            var cells = CellsByColumn(row);
            if (!IsSignificantRow(cells.Values, sharedStrings))
            {
                continue;
            }

            var item = metadata.Create();
            var rowIsValid = true;
            foreach (var pair in mapping)
            {
                cells.TryGetValue(pair.Key, out var cell);
                var value = cell is null ? null : ValueOf(cell, sharedStrings);
                var converter = pair.Value.Property.Converter;
                var conversionType = converter?.CellType ?? pair.Value.Property.Type;
                var isExcelError = cell?.DataType?.Value == CellValues.Error;
                var formulaWithoutCachedValue = cell?.CellFormula is not null && cell.CellValue is null;
                object? converted = null;
                var code = ExcelReadErrorCode.InvalidValue;
                var convertedSuccessfully = false;
                if (!isExcelError && !formulaWithoutCachedValue)
                {
                    convertedSuccessfully = CellValueConverter.TryConvert(
                        value,
                        conversionType,
                        cell is not null && IsDateCell(cell, dateStyles),
                        readOptions.Culture,
                        readOptions.EmptyStringAsNull,
                        out converted,
                        out code);
                }
                if (convertedSuccessfully && converted is not null && converter is not null)
                {
                    convertedSuccessfully = converter.TryRead(converted, out converted);
                    if (!convertedSuccessfully)
                    {
                        code = ExcelReadErrorCode.InvalidValue;
                    }
                }

                if (convertedSuccessfully)
                {
                    pair.Value.Property.SetValue(item, converted);

                    var declarativeValidation = pair.Value.Property.DeclarativeValidation;
                    if (declarativeValidation is not null && !declarativeValidation.IsValid(converted))
                    {
                        rowIsValid = false;
                        AddError(errors, new ExcelReadError(
                            sheetName,
                            rowNumber,
                            pair.Key,
                            CellReference(pair.Key, rowNumber),
                            pair.Value.Header,
                            pair.Value.Property.Name,
                            value,
                            pair.Value.Property.Type,
                            ExcelReadErrorCode.ValidationFailed,
                            declarativeValidation.Message), readOptions.MaxErrors);
                    }

                    foreach (var validation in pair.Value.Property.Validations)
                    {
                        if (validation.Predicate(converted))
                        {
                            continue;
                        }

                        rowIsValid = false;
                        AddError(errors, new ExcelReadError(
                            sheetName,
                            rowNumber,
                            pair.Key,
                            CellReference(pair.Key, rowNumber),
                            pair.Value.Header,
                            pair.Value.Property.Name,
                            value,
                            pair.Value.Property.Type,
                            ExcelReadErrorCode.ValidationFailed,
                            validation.Message), readOptions.MaxErrors);
                    }

                    continue;
                }

                rowIsValid = false;
                AddError(errors, new ExcelReadError(
                    sheetName,
                    rowNumber,
                    pair.Key,
                    CellReference(pair.Key, rowNumber),
                    pair.Value.Header,
                    pair.Value.Property.Name,
                    value,
                    pair.Value.Property.Type,
                    code,
                    ErrorMessage(
                        code,
                        pair.Value.Property.Name,
                        pair.Value.Property.Type,
                        value,
                        converter is not null,
                        isExcelError,
                        formulaWithoutCachedValue)), readOptions.MaxErrors);
            }

            if (rowIsValid || readOptions.InvalidRowPolicy == ExcelInvalidRowPolicy.Include)
            {
                successfulRows.Add(item);
            }
        }

        return new ExcelReadResult<T>(successfulRows, errors);
    }

    private static Dictionary<int, HeaderMapping> CreateHeaderMapping(
        Row? header,
        IReadOnlyList<ImportProperty> properties,
        HashSet<int> usedColumns,
        IReadOnlyList<string>? sharedStrings,
        string sheetName,
        List<ExcelReadError> errors,
        int maxErrors)
    {
        var mapping = new Dictionary<int, HeaderMapping>();
        var foundProperties = new HashSet<ImportProperty>();
        var seenHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var propertiesByHeader = properties.ToDictionary(property => property.Header, StringComparer.OrdinalIgnoreCase);
        var propertiesBySourceHeader = properties
            .Where(property => property.SourceHeader is not null)
            .ToDictionary(property => property.SourceHeader!, StringComparer.OrdinalIgnoreCase);

        foreach (var property in properties.Where(property => property.SourceColumnNumber is not null))
        {
            var columnNumber = property.SourceColumnNumber!.Value;
            if (usedColumns.Contains(columnNumber))
            {
                mapping.Add(columnNumber, new HeaderMapping(property, property.Header));
                foundProperties.Add(property);
            }
        }

        if (header is not null)
        {
            foreach (var pair in CellsByColumn(header))
            {
                var value = ValueOf(pair.Value, sharedStrings);
                if (string.IsNullOrWhiteSpace(value))
                {
                    continue;
                }

                var name = value!;

                if (!seenHeaders.Add(name))
                {
                    AddError(errors, new ExcelReadError(
                        sheetName,
                        header.RowIndex?.Value ?? 1U,
                        pair.Key,
                        CellReference(pair.Key, header.RowIndex?.Value ?? 1U),
                        name,
                        propertiesByHeader.TryGetValue(name, out var propertyForName) ? propertyForName.Name : null,
                        name,
                        propertiesByHeader.TryGetValue(name, out var duplicateProperty) ? duplicateProperty.Type : null,
                        ExcelReadErrorCode.DuplicateHeader,
                        $"Header '{name}' appears more than once."), maxErrors);
                    continue;
                }

                if (mapping.ContainsKey(pair.Key))
                {
                    continue;
                }

                var property = propertiesByHeader.TryGetValue(name, out var runtimeProperty)
                    ? runtimeProperty
                    : propertiesBySourceHeader.TryGetValue(name, out var schemaProperty)
                        ? schemaProperty
                        : null;
                if (property is not null && foundProperties.Add(property))
                {
                    mapping.Add(pair.Key, new HeaderMapping(property, name));
                }
            }
        }

        foreach (var property in properties.Where(property => property.IsRequired && !foundProperties.Contains(property)))
        {
            var identifier = property.SourceColumnNumber is not null
                ? $"Column {property.SourceColumnNumber.Value}"
                : property.SourceHeader ?? property.Header;
            AddError(errors, MissingColumnError(sheetName, property, identifier, header?.RowIndex?.Value ?? 1U), maxErrors);
        }

        return mapping;
    }

    private static ExcelReadError MissingColumnError(
        string sheetName,
        ImportProperty property,
        string identifier,
        uint rowNumber = 1U) => new(
            sheetName,
            rowNumber,
            null,
            null,
            property.Header,
            property.Name,
            null,
            property.Type,
            ExcelReadErrorCode.MissingHeader,
            $"Required column '{identifier}' for '{property.Header}' was not found.");

    private static WorksheetScan ScanRows(
        IEnumerable<Row> rows,
        IReadOnlyList<string>? sharedStrings,
        int maxRows,
        CancellationToken cancellationToken,
        IReadOnlyList<ImportProperty>? configuredProperties)
    {
        var columns = new HashSet<int>();
        Row? header = null;
        Row? firstNonEmptyRow = null;
        var rowCount = 0;
        foreach (var row in rows)
        {
            cancellationToken.ThrowIfCancellationRequested();
            rowCount++;
            if (rowCount > maxRows)
            {
                throw new InvalidDataException($"The worksheet exceeds the configured limit of {maxRows} physical rows.");
            }

            var cells = CellsByColumn(row);
            foreach (var pair in cells)
            {
                if (IsSignificantCell(pair.Value, sharedStrings))
                {
                    columns.Add(pair.Key);
                }
            }

            if (header is null && IsHeaderRow(cells.Values, sharedStrings))
            {
                firstNonEmptyRow ??= row;
                if (configuredProperties is null || MatchesConfiguredHeaders(cells.Values, sharedStrings, configuredProperties))
                {
                    header = row;
                }
            }
        }

        return new WorksheetScan(header ?? firstNonEmptyRow, columns);
    }

    private static bool MatchesConfiguredHeaders(
        IEnumerable<Cell> cells,
        IReadOnlyList<string>? sharedStrings,
        IReadOnlyList<ImportProperty> properties)
    {
        var requiredHeaders = properties.Where(property => property.IsRequired && property.SourceColumnNumber is null).ToArray();
        if (requiredHeaders.Length == 0)
        {
            // A schema that maps every field by physical position has no header identity to discover.
            return true;
        }

        var values = new HashSet<string>(
            cells.Select(cell => ValueOf(cell, sharedStrings)).Where(value => !string.IsNullOrWhiteSpace(value))!,
            StringComparer.OrdinalIgnoreCase);
        return requiredHeaders.All(property => values.Contains(property.Header)
            || property.SourceHeader is not null && values.Contains(property.SourceHeader));
    }

    private static bool IsHeaderRow(IEnumerable<Cell> cells, IReadOnlyList<string>? sharedStrings) => cells
        .Any(cell => !string.IsNullOrWhiteSpace(ValueOf(cell, sharedStrings)));

    private static bool IsSignificantRow(IEnumerable<Cell> cells, IReadOnlyList<string>? sharedStrings) => cells
        .Any(cell => IsSignificantCell(cell, sharedStrings));

    private static bool IsSignificantCell(Cell cell, IReadOnlyList<string>? sharedStrings) => cell.CellFormula is not null
        || cell.DataType?.Value == CellValues.Error
        || !string.IsNullOrEmpty(ValueOf(cell, sharedStrings));

    private static IEnumerable<Row> RowsAfter(IEnumerable<Row> rows, Row header)
    {
        var foundHeader = false;
        foreach (var row in rows)
        {
            if (!foundHeader)
            {
                foundHeader = ReferenceEquals(row, header);
                continue;
            }

            yield return row;
        }
    }

    private static Dictionary<int, Cell> CellsByColumn(Row row)
    {
        var cells = new Dictionary<int, Cell>();
        var nextColumn = 1;
        var rowIndex = row.RowIndex?.Value;
        if (rowIndex is not null && (rowIndex.Value < 1 || rowIndex.Value > MaximumExcelRow))
        {
            throw new InvalidDataException($"Worksheet row {rowIndex.Value} is outside Excel's valid row range.");
        }
        foreach (var cell in row.Elements<Cell>())
        {
            var coordinate = CellCoordinate(cell.CellReference?.Value);
            if (coordinate is not null && rowIndex is not null && coordinate.Value.row != rowIndex.Value)
            {
                throw new InvalidDataException($"Cell reference '{cell.CellReference!.Value}' does not match its containing row {rowIndex.Value}.");
            }

            var column = coordinate?.column ?? nextColumn;
            if (column > MaximumExcelColumn)
            {
                throw new InvalidDataException($"The worksheet contains a cell beyond Excel's maximum column {MaximumExcelColumn}.");
            }

            if (cells.ContainsKey(column))
            {
                throw new InvalidDataException($"The worksheet contains more than one cell in physical column {column} of the same row.");
            }

            cells.Add(column, cell);
            nextColumn = column + 1;
        }

        return cells;
    }

    private static (int column, uint row)? CellCoordinate(string? cellReference)
    {
        if (cellReference is null || cellReference.Length == 0)
        {
            return null;
        }

        var reference = cellReference;
        var column = 0;
        var split = 0;
        while (split < reference.Length && reference[split] is >= 'A' and <= 'Z')
        {
            var character = reference[split];
            var digit = character - 'A' + 1;
            if (column > (MaximumExcelColumn - digit) / 26)
            {
                throw new InvalidDataException($"Cell reference '{reference}' exceeds Excel's maximum column XFD.");
            }

            column = (column * 26) + digit;
            split++;
        }

        if (split == 0 || split == reference.Length
            || !uint.TryParse(reference.Substring(split), out var row)
            || row < 1 || row > MaximumExcelRow)
        {
            throw new InvalidDataException($"Cell reference '{reference}' is not a valid Excel A1 coordinate.");
        }

        return (column, row);
    }

    private static string? ValueOf(Cell cell, IReadOnlyList<string>? sharedStrings)
    {
        if (cell.DataType?.Value == CellValues.InlineString)
        {
            return cell.InlineString?.InnerText;
        }

        if (cell.DataType?.Value == CellValues.SharedString)
        {
            if (!int.TryParse(cell.CellValue?.Text, out var sharedStringIndex)
                || sharedStrings is null
                || sharedStringIndex < 0
                || sharedStringIndex >= sharedStrings.Count)
            {
                throw new InvalidDataException("The worksheet references a missing or invalid shared string.");
            }

            return sharedStrings[sharedStringIndex];
        }

        return cell.CellValue?.Text;
    }

    private static List<string>? SharedStrings(
        SharedStringTablePart? sharedStringPart,
        CancellationToken cancellationToken)
    {
        var table = sharedStringPart?.SharedStringTable;
        if (table is null)
        {
            return null;
        }

        var values = new List<string>();
        foreach (var item in table.Elements<SharedStringItem>())
        {
            cancellationToken.ThrowIfCancellationRequested();
            values.Add(item.InnerText);
        }

        return values;
    }

    private static HashSet<uint> DateStyles(WorkbookStylesPart? stylesPart, CancellationToken cancellationToken)
    {
        var dateFormats = new Dictionary<uint, string>();
        var formats = stylesPart?.Stylesheet?.NumberingFormats?.Elements<NumberingFormat>()
            ?? Enumerable.Empty<NumberingFormat>();
        foreach (var format in formats)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var numberFormatId = format.NumberFormatId?.Value;
            var formatCode = format.FormatCode?.Value;
            if (numberFormatId is not null && formatCode is not null)
            {
                dateFormats[numberFormatId.Value] = formatCode;
            }
        }

        var dateStyles = new HashSet<uint>();
        var styleIndex = 0U;
        var cellFormats = stylesPart?.Stylesheet?.CellFormats?.Elements<CellFormat>()
            ?? Enumerable.Empty<CellFormat>();
        foreach (var cellFormat in cellFormats)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var numberFormatId = cellFormat.NumberFormatId?.Value;
            if (numberFormatId is not null
                && (BuiltInDateFormats.Contains(numberFormatId.Value)
                    || dateFormats.TryGetValue(numberFormatId.Value, out var format) && HasDateParts(format)))
            {
                dateStyles.Add(styleIndex);
            }

            styleIndex++;
        }

        return dateStyles;
    }

    private static bool IsDateCell(Cell cell, HashSet<uint> dateStyles) => cell.StyleIndex?.Value is uint styleIndex
        && dateStyles.Contains(styleIndex);

    private static bool HasDateParts(string format)
    {
#if NET8_0_OR_GREATER
        return format.Contains('y') || format.Contains('d');
#else
        return format.IndexOf('y') >= 0 || format.IndexOf('d') >= 0;
#endif
    }

    private static string CellReference(int columnNumber, uint rowNumber)
    {
        var column = string.Empty;
        var value = columnNumber;
        while (value > 0)
        {
            value--;
            column = (char)('A' + (value % 26)) + column;
            value /= 26;
        }

        return column + rowNumber.ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    private static string ErrorMessage(
        ExcelReadErrorCode code,
        string propertyName,
        Type propertyType,
        string? value,
        bool customConverter,
        bool isExcelError,
        bool formulaWithoutCachedValue)
    {
        if (isExcelError)
        {
            return $"Cell contains the Excel error value '{value}'.";
        }

        if (formulaWithoutCachedValue)
        {
            return "Formula cells require a cached value to be imported.";
        }

        return code == ExcelReadErrorCode.RequiredValueMissing
            ? $"A value is required for '{propertyName}'."
            : customConverter
                ? $"Value '{value}' could not be converted to '{propertyType.FullName}'."
            : $"Value cannot be converted to '{propertyType.FullName}' for '{propertyName}'.";
    }

    private static void AddError(List<ExcelReadError> errors, ExcelReadError error, int maxErrors)
    {
        if (errors.Count >= maxErrors)
        {
            throw new InvalidDataException($"The worksheet exceeds the configured limit of {maxErrors} data errors.");
        }

        errors.Add(error);
    }

    private sealed class WorksheetScan
    {
        internal WorksheetScan(Row? header, HashSet<int> usedColumns)
        {
            Header = header;
            UsedColumns = usedColumns;
        }

        internal Row? Header { get; }

        internal HashSet<int> UsedColumns { get; }
    }

    private sealed class HeaderMapping
    {
        internal HeaderMapping(ImportProperty property, string header)
        {
            Property = property;
            Header = header;
        }

        internal ImportProperty Property { get; }

        internal string Header { get; }
    }
}
