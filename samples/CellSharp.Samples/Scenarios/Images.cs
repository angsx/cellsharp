using CellSharp;

namespace CellSharp.Samples.Scenarios;

internal static class Images
{
    private static readonly Uri PlaceholderImage = new("https://placehold.co/160x80/png?text=CellSharp", UriKind.Absolute);
    private const string PlaceholderHost = "placehold.co";
    private const int MaximumPlaceholderBytes = 1024 * 1024;

    internal static void Run(string outputDirectory)
    {
        var rows = new[] { new SalesRow("North", 12000m), new SalesRow("South", 6200m) };
        var schema = Excel.Schema<SalesRow>()
            .SheetName("Sales")
            .AsTable("BrandedSales")
            .Column(row => row.Region)
            .Column(row => row.Revenue, column => column.Format("€ #,##0"))
            .Build();
        using var logo = new MemoryStream(DownloadPlaceholder());
        Excel.Workbook().AddSheet(rows, schema, sheet =>
        {
            sheet.DataStartAt("A7");
            sheet.AddImage(logo, ExcelImageFormat.Png).At("A1", offsetX: 6, offsetY: 6).Width(96).KeepAspectRatio().Name("CellSharp logo").AltText("CellSharp logo");
            sheet.Merge("B1:E2").Value("Branded sales report").Style(s => s.Bold().FontSize(18).FillColor("#1F4E78").FontColor("#FFFFFF").AlignCenter().VerticalAlignCenter());
            sheet.Range("B8:B200").ConditionalFormat().GreaterThan(10000).Style(s => s.FillColor("#C6EFCE").Bold());
            sheet.Footer.Right("Page &P of &N");
            sheet.Column("A").Width(14);
            sheet.Columns("B:E").Width(14);
        }).Write(Path.Combine(outputDirectory, "Images.xlsx"));
    }

    private static byte[] DownloadPlaceholder()
    {
        if (PlaceholderImage.Scheme != Uri.UriSchemeHttps || !string.Equals(PlaceholderImage.Host, PlaceholderHost, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("The sample placeholder URL must use HTTPS and the approved placeholder host.");
        }

        using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(15) };
        var bytes = client.GetByteArrayAsync(PlaceholderImage).GetAwaiter().GetResult();
        if (bytes.Length == 0 || bytes.Length > MaximumPlaceholderBytes)
        {
            throw new InvalidDataException("The downloaded placeholder image has an invalid size.");
        }

        return bytes;
    }

    private sealed record SalesRow(string Region, decimal Revenue);
}
