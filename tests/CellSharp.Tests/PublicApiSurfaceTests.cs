using System.Reflection;
using Xunit;

namespace CellSharp.Tests;

public sealed class PublicApiSurfaceTests
{
    private static readonly Type[] ExpectedTypes =
    [
        typeof(Excel), typeof(ExcelBorderBuilder), typeof(ExcelBorderStyle), typeof(ExcelCellHorizontalAlignment),
        typeof(ExcelColumnAttribute), typeof(ExcelColumnBuilder<,>), typeof(ExcelColumn), typeof(ExcelColumns),
        typeof(ExcelConditionalFormatBuilder), typeof(ExcelConditionalFormatRuleBuilder), typeof(ExcelFormulaContext),
        typeof(ExcelHeaderFooter), typeof(ExcelHeaderStyleBuilder), typeof(ExcelHorizontalAlignment), typeof(ExcelImage),
        typeof(ExcelImageFormat), typeof(ExcelIgnoreAttribute), typeof(ExcelInvalidRowPolicy), typeof(ExcelKpi),
        typeof(ExcelRange), typeof(ExcelReadError), typeof(ExcelReadErrorCode), typeof(ExcelReadErrorKind),
        typeof(ExcelReadOptions), typeof(ExcelReadResult<>), typeof(ExcelRow), typeof(ExcelSchema<>),
        typeof(ExcelSchemaBuilder<>), typeof(ExcelSchemaOverlay<>), typeof(ExcelSchemaOverlayBuilder<>), typeof(ExcelStyle),
        typeof(ExcelStyleBuilder), typeof(ExcelStyleTemplate), typeof(ExcelTheme), typeof(ExcelUnderline),
        typeof(ExcelVerticalAlignment), typeof(ExcelWorkbookBuilder), typeof(ExcelWorkbookReader), typeof(ExcelWorksheetBuilder),
        typeof(ExcelWriteOptionsBuilder), typeof(IExcelValueConverter<,>),
    ];

    [Fact]
    public void ExportedTypeInventoryIsDeliberate()
    {
        var actual = typeof(Excel).Assembly.GetExportedTypes().OrderBy(type => type.FullName).ToArray();
        var expected = ExpectedTypes.OrderBy(type => type.FullName).ToArray();
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void PublicSignaturesDoNotLeakOpenXmlOrCellSharpInternalTypes()
    {
        var signatures = typeof(Excel).Assembly.GetExportedTypes()
            .SelectMany(PublicSignatureTypes)
            .SelectMany(Flatten)
            .Where(type => type.Namespace is not null)
            .ToArray();

        Assert.DoesNotContain(signatures, type => type.Namespace!.StartsWith("DocumentFormat.OpenXml", StringComparison.Ordinal));
        Assert.DoesNotContain(signatures, type => type.Namespace!.StartsWith("CellSharp.Internal", StringComparison.Ordinal));
    }

    private static IEnumerable<Type> PublicSignatureTypes(Type type)
    {
        yield return type;
        foreach (var constructor in type.GetConstructors())
            foreach (var parameter in constructor.GetParameters()) yield return parameter.ParameterType;
        foreach (var method in type.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly))
        {
            yield return method.ReturnType;
            foreach (var parameter in method.GetParameters()) yield return parameter.ParameterType;
        }
        foreach (var property in type.GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly)) yield return property.PropertyType;
    }

    private static IEnumerable<Type> Flatten(Type type)
    {
        yield return type;
        if (type.IsByRef || type.IsPointer || type.IsArray)
        {
            foreach (var item in Flatten(type.GetElementType()!)) yield return item;
        }
        foreach (var argument in type.IsGenericType ? type.GetGenericArguments() : Type.EmptyTypes)
            foreach (var item in Flatten(argument)) yield return item;
    }
}
