using System.Security;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CellSharp.Internal;

internal static class XlsxCommentsWriter
{
    internal static void Apply(IReadOnlyList<WorksheetValidationContext> worksheets)
    {
        foreach (var context in worksheets)
        {
            var comments = context.Layout?.Comments;
            if (comments is null || comments.Count == 0) continue;
            var authors = comments.Values.Select(comment => comment.Author).Distinct(StringComparer.Ordinal).ToArray();
            var authorIds = authors.Select((author, index) => new { author, index }).ToDictionary(value => value.author, value => value.index, StringComparer.Ordinal);
            var commentsPart = context.WorksheetPart.AddNewPart<WorksheetCommentsPart>();
            var commentList = new CommentList();
            foreach (var item in comments.OrderBy(value => value.Key.FromRow).ThenBy(value => value.Key.FromColumn))
            {
                var comment = new Comment { Reference = item.Key.ToString(), AuthorId = (uint)authorIds[item.Value.Author] };
                comment.AppendChild(new CommentText(new Run(new Text(item.Value.Text) { Space = SpaceProcessingModeValues.Preserve })));
                commentList.AppendChild(comment);
            }
            commentsPart.Comments = new Comments(new Authors(authors.Select(author => new Author(author))), commentList);

            var vmlPart = context.WorksheetPart.AddNewPart<VmlDrawingPart>();
            WriteVml(vmlPart, comments.Keys.OrderBy(range => range.FromRow).ThenBy(range => range.FromColumn));
            context.Worksheet.AppendChild(new LegacyDrawing { Id = context.WorksheetPart.GetIdOfPart(vmlPart) });
        }
    }

    private static void WriteVml(VmlDrawingPart part, IEnumerable<ExcelRangeReference> ranges)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream, new UTF8Encoding(false));
        writer.Write("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
        writer.Write("<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/><v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype>");
        var shapeId = 1025;
        foreach (var range in ranges)
        {
            writer.Write($"<v:shape id=\"_x0000_s{shapeId++}\" type=\"#_x0000_t202\" style=\"position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden\" fillcolor=\"#ffffe1\" o:insetmode=\"auto\"><v:fill color2=\"#ffffe1\"/><v:shadow on=\"t\" color=\"black\" obscured=\"t\"/><v:path o:connecttype=\"none\"/><v:textbox style=\"mso-direction-alt:auto\"><div style=\"text-align:left\"/></v:textbox><x:ClientData ObjectType=\"Note\"><x:MoveWithCells/><x:SizeWithCells/><x:AutoFill>False</x:AutoFill><x:Row>{range.FromRow - 1}</x:Row><x:Column>{range.FromColumn - 1}</x:Column></x:ClientData></v:shape>");
        }
        writer.Write("</xml>");
    }
}
