using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CellSharp.Internal;

internal static class XlsxDataValidationWriter
{
    private const string DefaultLookupSheetName = "_CellSharpValidation";
    private const uint MaximumExcelRow = 1048576U;

    internal static void Apply(WorkbookPart workbookPart, IReadOnlyList<WorksheetValidationContext> worksheets, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var validated = worksheets
            .SelectMany(worksheet => worksheet.Properties.Select((property, index) => new ValidatedColumn(
                worksheet.Worksheet,
                index + worksheet.DataStartColumn - 1,
                property.DeclarativeValidation)))
            .Where(column => column.Validation is not null)
            .ToArray();
        if (validated.Length == 0)
        {
            return;
        }

        var listFormulas = CreateLookupSheet(workbookPart, validated, LookupSheetName(workbookPart), cancellationToken);
        foreach (var worksheet in worksheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var worksheetColumns = validated.Where(column => column.Worksheet == worksheet.Worksheet).ToArray();
            if (worksheetColumns.Length == 0)
            {
                continue;
            }

            var validations = new DataValidations();
            foreach (var column in worksheetColumns)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var validation = new DataValidation
                {
                    AllowBlank = true,
                    ShowErrorMessage = true,
                    ShowInputMessage = true,
                SequenceOfReferences = new ListValue<StringValue> { InnerText = Range(column.Index, worksheets.First(context => context.Worksheet == column.Worksheet).DataStartRow) },
                };
                switch (column.Validation)
                {
                    case AllowedValuesValidationRule allowed:
                        validation.Type = DataValidationValues.List;
                        validation.Append(new Formula1(listFormulas[allowed]));
                        break;
                    case NumericRangeValidationRule range:
                        validation.Type = DataValidationValues.Decimal;
                        validation.Operator = DataValidationOperatorValues.Between;
                        validation.Append(
                            new Formula1(Number(range.Minimum)),
                            new Formula2(Number(range.Maximum)));
                        break;
                    case DateRangeValidationRule dates:
                        validation.Type = DataValidationValues.Date;
                        validation.Operator = DataValidationOperatorValues.Between;
                        validation.Append(
                            new Formula1(Number(dates.Minimum.ToOADate())),
                            new Formula2(Number(dates.Maximum.ToOADate())));
                        break;
                    default:
                        throw new InvalidOperationException("The declarative validation rule is not supported by XLSX export.");
                }

                validations.Append(validation);
            }

            validations.Count = (uint)worksheetColumns.Length;
            var next = worksheet.Worksheet.GetFirstChild<Hyperlinks>();
            if (next is null) worksheet.Worksheet.AppendChild(validations);
            else worksheet.Worksheet.InsertBefore(validations, next);
        }
    }

    internal static bool IsInternalSheet(Sheet sheet) => sheet.State?.Value == SheetStateValues.Hidden
        && (sheet.Name?.Value?.StartsWith(DefaultLookupSheetName, StringComparison.OrdinalIgnoreCase) ?? false);

    private static Dictionary<AllowedValuesValidationRule, string> CreateLookupSheet(
        WorkbookPart workbookPart,
        IReadOnlyList<ValidatedColumn> validated,
        string lookupSheetName,
        CancellationToken cancellationToken)
    {
        var lists = validated
            .Select(column => column.Validation)
            .OfType<AllowedValuesValidationRule>()
            .Distinct(AllowedValuesComparer.Instance)
            .ToArray();
        var formulas = new Dictionary<AllowedValuesValidationRule, string>(AllowedValuesComparer.Instance);
        if (lists.Length == 0)
        {
            return formulas;
        }

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        var maximumValues = lists.Max(list => list.Values.Count);
        var rows = Enumerable.Range(1, maximumValues)
            .Select(rowNumber => new Row { RowIndex = (uint)rowNumber })
            .ToArray();
        for (var listIndex = 0; listIndex < lists.Length; listIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var column = ColumnName(listIndex + 1);
            var list = lists[listIndex];
            for (var valueIndex = 0; valueIndex < list.Values.Count; valueIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var cell = CellValueWriter.Create(list.Values[valueIndex]);
                cell.CellReference = $"{column}{valueIndex + 1}";
                rows[valueIndex].AppendChild(cell);
            }

            formulas.Add(list, $"'{lookupSheetName}'!${column}$1:${column}${list.Values.Count}");
        }

        foreach (var row in rows)
        {
            cancellationToken.ThrowIfCancellationRequested();
            sheetData.AppendChild(row);
        }

        worksheetPart.Worksheet = new Worksheet(sheetData);
        var workbook = workbookPart.Workbook
            ?? throw new InvalidOperationException("The workbook part has no workbook.");
        var sheets = workbook.Sheets!;
        var sheetId = sheets.Elements<Sheet>().Select(sheet => sheet.SheetId?.Value ?? 0U).DefaultIfEmpty(0U).Max() + 1U;
        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = lookupSheetName,
            State = SheetStateValues.Hidden,
        });
        worksheetPart.Worksheet.Save();
        return formulas;
    }

    private static string LookupSheetName(WorkbookPart workbookPart)
    {
        var usedNames = new HashSet<string>(
            (workbookPart.Workbook
                ?? throw new InvalidOperationException("The workbook part has no workbook."))
                .Sheets!.Elements<Sheet>()
                .Select(sheet => sheet.Name?.Value ?? string.Empty),
            StringComparer.OrdinalIgnoreCase);
        var name = DefaultLookupSheetName;
        var suffix = 2;
        while (usedNames.Contains(name))
        {
            name = DefaultLookupSheetName + suffix;
            suffix++;
        }

        return name;
    }

    private static string Range(int columnIndex, int dataStartRow) => $"{ColumnName(columnIndex + 1)}{dataStartRow + 1}:{ColumnName(columnIndex + 1)}{MaximumExcelRow}";

    private static string ColumnName(int columnNumber)
    {
        var name = string.Empty;
        var value = columnNumber;
        while (value > 0)
        {
            value--;
            name = (char)('A' + (value % 26)) + name;
            value /= 26;
        }

        return name;
    }

    private static string Number(object value) => Convert.ToString(value, CultureInfo.InvariantCulture)
        ?? throw new InvalidOperationException("A numeric validation bound cannot be null.");

    private sealed class ValidatedColumn
    {
        internal ValidatedColumn(Worksheet worksheet, int index, DeclarativeValidationRule? validation)
        {
            Worksheet = worksheet;
            Index = index;
            Validation = validation;
        }

        internal Worksheet Worksheet { get; }

        internal int Index { get; }

        internal DeclarativeValidationRule? Validation { get; }
    }

    private sealed class AllowedValuesComparer : IEqualityComparer<AllowedValuesValidationRule>
    {
        internal static AllowedValuesComparer Instance { get; } = new();

        public bool Equals(AllowedValuesValidationRule? x, AllowedValuesValidationRule? y) => ReferenceEquals(x, y)
            || (x is not null && y is not null && x.Values.SequenceEqual(y.Values, StringComparer.Ordinal));

        public int GetHashCode(AllowedValuesValidationRule value)
        {
            var hash = 17;
            foreach (var item in value.Values)
            {
                unchecked
                {
                    hash = (hash * 31) + StringComparer.Ordinal.GetHashCode(item);
                }
            }

            return hash;
        }
    }
}
