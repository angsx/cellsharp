using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelReleaseHardeningTests
{
    [Fact]
    public void CombinedWorkbookRemainsStructurallyValidWithOverlayTableFormulaValidationAndPrintSettings()
    {
        var path = Path.Combine(Path.GetTempPath(), $"cellsharp-hardening-{Guid.NewGuid():N}.xlsx");
        var schema = Excel.Schema<Order>()
            .SheetName("Order Details")
            .AsTable("OrdersTable")
            .AutoFilter()
            .FreezePanes(1, 1)
            .Landscape()
            .FitToPage(1, 0)
            .RepeatHeaderRowOnPrint()
            .PrintGridlines()
            .Column(order => order.Id)
            .Column(order => order.Category, column => column.AllowedValues("Retail", "Wholesale"))
            .Column(order => order.Quantity)
            .Column(order => order.Total, column => column.Formula(context => $"=A{context.Row}*C{context.Row}").Format("#,##0.00"))
            .Build();
        var overlay = Excel.Overlay<Order>(runtime => runtime.Include(order => order.Quantity, false));

        try
        {
            Excel.Workbook()
                .AddSheet([new Order { Id = 7, Category = "Retail", Quantity = 2 }], schema, overlay)
                .Write(path);

            using var document = SpreadsheetDocument.Open(path, false);
            var sheets = document.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().ToArray();
            var sheet = sheets.Single(item => item.Name!.Value == "Order Details");
            var worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id!);
            var table = Assert.Single(worksheetPart.TableDefinitionParts).Table!;
            var formula = worksheetPart.Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().Skip(1).Single().Elements<Cell>().ElementAt(2);
            var title = Assert.Single(document.WorkbookPart.Workbook!.DefinedNames!.Elements<DefinedName>());

            Assert.Equal("A1:C2", table.Reference!.Value);
            Assert.Equal(["Id", "Category", "Total"], table.TableColumns!.Elements<TableColumn>().Select(column => column.Name!.Value));
            Assert.Null(worksheetPart.Worksheet!.GetFirstChild<AutoFilter>());
            Assert.NotNull(worksheetPart.Worksheet!.GetFirstChild<DataValidations>());
            Assert.Equal("A2*C2", formula.CellFormula!.InnerText);
            Assert.Equal("'Order Details'!$1:$1", title.Text);
            Assert.NotNull(document.WorkbookPart.Workbook!.CalculationProperties);
            Assert.Single(sheets, item => item.State?.Value == SheetStateValues.Hidden);
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    private sealed class Order
    {
        public int Id { get; set; }

        public string? Category { get; set; }

        public int Quantity { get; set; }

        public decimal Total { get; set; }
    }
}
