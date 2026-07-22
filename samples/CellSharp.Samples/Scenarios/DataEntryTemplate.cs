using CellSharp;
using CellSharp.Samples.Models;

namespace CellSharp.Samples.Scenarios;

internal static class DataEntryTemplate
{
    internal static ExcelSchema<ProductImportRow> Schema { get; } = Excel.Schema<ProductImportRow>()
        .SheetName("Product import")
        .FreezePanes(1, 0)
        .Column(x => x.Name, column => column.Header("Product name").Width(30))
        .Column(x => x.Category, column => column.Header("Category").Optional().AllowedValues("Hardware", "Software", "Service").Width(16))
        .Column(x => x.Quantity, column => column.Header("Opening quantity").Range(0, 100000).Format("0").Width(18))
        .Column(x => x.AvailableFrom, column => column.Header("Available from").DateBetween(new DateTime(2025, 1, 1), new DateTime(2030, 12, 31)).Format("yyyy-mm-dd").Width(16))
        .Column(x => x.Notes, column => column.Header("Notes").Optional().Width(36))
        .Build();

    internal static void Run(string outputDirectory)
    {
        Excel.CreateTemplate(
            Path.Combine(outputDirectory, "product-import-template.xlsx"),
            Schema,
            options => options.Theme(ExcelTheme.Modern).FreezeHeaderRow());
    }
}
