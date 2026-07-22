using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CellSharp.Internal;

internal static class XlsxTableWriter
{
    internal static void Apply(IReadOnlyList<WorksheetValidationContext> worksheets, CancellationToken cancellationToken = default)
    {
        var tableWorksheets = worksheets.Where(worksheet => worksheet.Table is not null).ToArray();
        if (tableWorksheets.Length == 0)
        {
            return;
        }

        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var index = 0; index < tableWorksheets.Length; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var worksheet = tableWorksheets[index];
            var name = ResolveName(worksheet, names);
            names.Add(name);
            AddTable(worksheet, (uint)(index + 1), name);
        }
    }

    private static void AddTable(WorksheetValidationContext worksheet, uint tableId, string name)
    {
        var reference = $"{ColumnName(worksheet.DataStartColumn)}{worksheet.DataStartRow}:{ColumnName(worksheet.DataStartColumn + worksheet.Properties.Count - 1)}{worksheet.DataStartRow + worksheet.DataRowCount}";
        var tablePart = worksheet.WorksheetPart.AddNewPart<TableDefinitionPart>();
        var columns = new TableColumns { Count = (uint)worksheet.Properties.Count };
        for (var index = 0; index < worksheet.Properties.Count; index++)
        {
            columns.AppendChild(new TableColumn
            {
                Id = (uint)(index + 1),
                Name = worksheet.Properties[index].Header,
            });
        }

        var table = new Table
        {
            Id = tableId,
            Name = name,
            DisplayName = name,
            Reference = reference,
            AutoFilter = new AutoFilter { Reference = reference },
            TableColumns = columns,
        };
        if (worksheet.Table!.Style is not null)
        {
            table.TableStyleInfo = new TableStyleInfo
            {
                Name = worksheet.Table.Style,
                ShowRowStripes = true,
                ShowColumnStripes = false,
            };
        }

        tablePart.Table = table;
        var tableParts = worksheet.Worksheet.GetFirstChild<TableParts>() ?? worksheet.Worksheet.AppendChild(new TableParts());
        tableParts.AppendChild(new TablePart { Id = worksheet.WorksheetPart.GetIdOfPart(tablePart) });
        tableParts.Count = (uint)tableParts.Elements<TablePart>().Count();
    }

    private static string ResolveName(WorksheetValidationContext worksheet, HashSet<string> usedNames)
    {
        var explicitName = worksheet.Table!.Name;
        if (explicitName is not null)
        {
            if (usedNames.Contains(explicitName))
            {
                throw new InvalidOperationException($"Table name '{explicitName}' is already used in this workbook.");
            }

            return explicitName;
        }

        var baseName = AutomaticName(worksheet);
        var name = baseName;
        var suffix = 2;
        while (usedNames.Contains(name))
        {
            name = baseName + suffix;
            suffix++;
        }

        return name;
    }

    private static string AutomaticName(WorksheetValidationContext worksheet)
    {
        var source = worksheet.SheetName;
        var normalized = new string(source.Where(character => character is >= 'A' and <= 'Z'
            || character is >= 'a' and <= 'z'
            || character is >= '0' and <= '9'
            || character == '_').ToArray());
        if (string.IsNullOrEmpty(normalized) || !char.IsLetter(normalized[0]) && normalized[0] != '_')
        {
            normalized = "Table" + normalized;
        }

        return normalized + "Table";
    }

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
}
