using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

namespace CellSharp.Internal;

internal static class XlsxDefinedNamesWriter
{
    internal static void Apply(WorkbookPart workbookPart, IReadOnlyList<WorksheetValidationContext> worksheets)
    {
        var definitions = worksheets.SelectMany(worksheet => (worksheet.Layout?.DefinedNames ?? Enumerable.Empty<DefinedNameDefinition>())
            .Select(definition => (worksheet, definition))).ToArray();
        if (definitions.Length == 0) return;
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var definition in definitions)
            if (!names.Add(definition.definition.Name)) throw new InvalidOperationException($"A defined name named '{definition.definition.Name}' already exists.");
        var workbook = workbookPart.Workbook ?? throw new InvalidOperationException("The workbook part has no workbook.");
        var definedNames = workbook.DefinedNames ?? workbook.AppendChild(new DefinedNames());
        foreach (var item in definitions)
            definedNames.AppendChild(new DefinedName { Name = item.definition.Name, Text = ExcelDefinedName.AbsoluteReference(item.worksheet.SheetName, item.definition.Range) });
    }
}
