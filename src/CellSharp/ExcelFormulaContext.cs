namespace CellSharp;

/// <summary>Provides the real one-based Excel coordinates for one exported formula cell.</summary>
public sealed class ExcelFormulaContext
{
    internal ExcelFormulaContext(uint row, int column, string sheetName)
    {
        Row = row;
        Column = column;
        SheetName = sheetName;
    }

    /// <summary>Gets the one-based Excel row number. The first data row is 2.</summary>
    public uint Row { get; }

    /// <summary>Gets the one-based Excel column number.</summary>
    public int Column { get; }

    /// <summary>Gets the worksheet name.</summary>
    public string SheetName { get; }
}
