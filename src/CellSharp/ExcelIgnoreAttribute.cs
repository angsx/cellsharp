namespace CellSharp;

/// <summary>Excludes a property from an attribute-based schema.</summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
public sealed class ExcelIgnoreAttribute : Attribute
{
}
