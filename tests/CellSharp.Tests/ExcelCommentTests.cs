using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelCommentTests
{
    private static readonly string[] SalesAuthors = ["CellSharp", "Finance"];
    [Fact]
    public void ClassicCommentsDeduplicateAuthorsAndCoexistWithStylesHyperlinksAndMerges()
    {
        var path = TemporaryPath();
        try
        {
            Excel.Workbook()
                .AddSheet("Sales", sheet =>
                {
                    sheet.Merge("A1:B1").Value("Sales");
                    sheet.Cell("A1").Comment("Top-left merge comment", "CellSharp");
                    sheet.Cell("C3").Value("OpenAI").Hyperlink("https://openai.com").Style(s => s.Bold()).Comment("External source", "CellSharp");
                    sheet.Cell("D4").Comment("Finance note", "Finance");
                })
                .AddSheet("Costs", sheet => sheet.Cell("A1").Comment("Cost note"))
                .Write(path);

            using var document = SpreadsheetDocument.Open(path, false);
            var sales = document.WorkbookPart!.WorksheetParts.First();
            var comments = sales.WorksheetCommentsPart!.Comments!;
            Assert.Equal(SalesAuthors, comments.Authors!.Elements<Author>().Select(author => author.Text));
            Assert.Equal(3, comments.CommentList!.Elements<Comment>().Count());
            Assert.Single(sales.VmlDrawingParts);
            Assert.NotNull(sales.Worksheet!.GetFirstChild<LegacyDrawing>());
            Assert.NotNull(sales.Worksheet.GetFirstChild<Hyperlinks>());
            Assert.Equal(2, document.WorkbookPart.WorksheetParts.Count(part => part.WorksheetCommentsPart is not null));
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
        finally { Delete(path); }
    }

    [Fact]
    public void EmptyAndDuplicateCommentsFailFast()
    {
        Assert.Throws<ArgumentException>(() => Excel.Workbook().AddSheet("Sales", sheet => sheet.Cell("A1").Comment(" ")));
        Assert.Throws<InvalidOperationException>(() => Excel.Workbook().AddSheet("Sales", sheet =>
        {
            sheet.Cell("A1").Comment("First");
            sheet.Cell("A1").Comment("Second");
        }));
        Assert.Throws<InvalidOperationException>(() => Excel.Workbook().AddSheet("Sales", sheet =>
        {
            sheet.Merge("A1:B1");
            sheet.Cell("B1").Comment("Not the merge anchor");
        }));
    }

    private static string TemporaryPath() => Path.Combine(Path.GetTempPath(), $"cellsharp-comment-{Guid.NewGuid():N}.xlsx");
    private static void Delete(string path) { if (File.Exists(path)) File.Delete(path); }
}
