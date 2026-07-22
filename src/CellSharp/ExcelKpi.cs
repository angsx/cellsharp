namespace CellSharp;

/// <summary>Exposes the ordinary worksheet ranges created by a KPI convenience component.</summary>
public sealed class ExcelKpi
{
    internal ExcelKpi(ExcelRange range, ExcelRange label, ExcelRange value) { Range = range; Label = label; Value = value; }
    /// <summary>Gets the complete KPI container range.</summary>
    public ExcelRange Range { get; }
    /// <summary>Gets the label area, normally the KPI's top row.</summary>
    public ExcelRange Label { get; }
    /// <summary>Gets the value area below the label. Use normal range APIs for formulas, formats, styles, names, or conditional formatting.</summary>
    public ExcelRange Value { get; }
}
