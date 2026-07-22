using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CellSharp.Internal;

internal static class CellValueWriter
{
    internal static Cell Create(object? value, uint? styleIndex = null)
    {
        if (value is null)
        {
            return Styled(new Cell(), styleIndex);
        }

        var cell = value switch
        {
            string text => TextCell(text),
            bool boolean => new Cell { DataType = CellValues.Boolean, CellValue = new CellValue(boolean ? "1" : "0") },
            DateTime dateTime => new Cell
            {
                DataType = CellValues.Number,
                CellValue = new CellValue(dateTime.ToOADate().ToString(CultureInfo.InvariantCulture)),
            },
            Guid guid => TextCell(guid.ToString("D")),
            byte number => NumberCell(number),
            sbyte number => NumberCell(number),
            short number => NumberCell(number),
            ushort number => NumberCell(number),
            int number => NumberCell(number),
            uint number => NumberCell(number),
            long number => NumberCell(number),
            ulong number => NumberCell(number),
            float number => NumberCell(number),
            double number => NumberCell(number),
            decimal number => NumberCell(number),
            _ => throw new NotSupportedException($"Values of type '{value.GetType().FullName}' cannot be exported."),
        };

        return Styled(cell, styleIndex);
    }

    internal static int DisplayWidth(object? value)
    {
        if (value is null)
        {
            return 0;
        }

        return value switch
        {
            DateTime dateTime => dateTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture).Length,
            Guid guid => guid.ToString("D").Length,
            IFormattable formattable => formattable.ToString(null, CultureInfo.InvariantCulture).Length,
            _ => value.ToString()?.Length ?? 0,
        };
    }

    internal static Cell CreateFormula(string formula, object? cachedValue, uint? styleIndex = null)
    {
        var cell = cachedValue switch
        {
            null => new Cell(),
            string text => new Cell { DataType = CellValues.String, CellValue = new CellValue(text) },
            bool boolean => new Cell { DataType = CellValues.Boolean, CellValue = new CellValue(boolean ? "1" : "0") },
            DateTime dateTime => new Cell
            {
                DataType = CellValues.Number,
                CellValue = new CellValue(dateTime.ToOADate().ToString(CultureInfo.InvariantCulture)),
            },
            Guid guid => new Cell { DataType = CellValues.String, CellValue = new CellValue(guid.ToString("D")) },
            byte number => NumberCell(number),
            sbyte number => NumberCell(number),
            short number => NumberCell(number),
            ushort number => NumberCell(number),
            int number => NumberCell(number),
            uint number => NumberCell(number),
            long number => NumberCell(number),
            ulong number => NumberCell(number),
            float number => NumberCell(number),
            double number => NumberCell(number),
            decimal number => NumberCell(number),
            _ => throw new NotSupportedException($"Values of type '{cachedValue.GetType().FullName}' cannot be exported."),
        };

        cell.CellFormula = new CellFormula(formula);
        return Styled(cell, styleIndex);
    }

    private static Cell NumberCell<T>(T value)
        where T : IFormattable => new()
        {
            DataType = CellValues.Number,
            CellValue = new CellValue(value.ToString(null, CultureInfo.InvariantCulture)),
        };

    private static Cell TextCell(string value) => new()
    {
        DataType = CellValues.InlineString,
        InlineString = new InlineString(new Text(value) { Space = SpaceProcessingModeValues.Preserve }),
    };

    private static Cell Styled(Cell cell, uint? styleIndex)
    {
        if (styleIndex is not null)
        {
            cell.StyleIndex = styleIndex.Value;
        }

        return cell;
    }
}
