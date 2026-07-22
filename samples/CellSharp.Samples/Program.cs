using CellSharp.Samples;
using CellSharp.Samples.Scenarios;

var outputDirectory = SampleOutput.Prepare();

Console.WriteLine("CellSharp Samples");
Console.WriteLine();

var scenarios = new (string Name, Action<string> Run)[]
{
    ("Basic export", BasicExport.Run),
    ("Basic import", BasicImport.Run),
    ("Fluent schema", FluentSchema.Run),
    ("Attribute schema", AttributeSchema.Run),
    ("Custom converter", CustomConverter.Run),
    ("Data-entry template", DataEntryTemplate.Run),
    ("Import round trip", ImportRoundTrip.Run),
    ("Multi-sheet workbook", MultiSheetWorkbook.Run),
    ("Formula columns", FormulaColumns.Run),
    ("Report and table", Report.Run),
    ("Conditional formatting", ConditionalFormatting.Run),
    ("Workbook utilities", WorkbookUtilities.Run),
    ("Images", Images.Run),
    ("Sales report", SalesReport.Run),
    ("Streams", Streams.Run),
};

foreach (var scenario in scenarios)
{
    scenario.Run(outputDirectory);
    Console.WriteLine($"✓ {scenario.Name}");
}

Console.WriteLine();
Console.WriteLine("Generated workbooks:");
foreach (var file in Directory.GetFiles(outputDirectory, "*.xlsx").OrderBy(Path.GetFileName))
{
    Console.WriteLine($"  {Path.GetRelativePath(Directory.GetCurrentDirectory(), file)}");
}

Console.WriteLine();
Console.WriteLine("Done. Open the generated workbooks to inspect the results.");
