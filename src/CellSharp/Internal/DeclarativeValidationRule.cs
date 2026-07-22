using System.Globalization;

namespace CellSharp.Internal;

internal abstract class DeclarativeValidationRule
{
    internal abstract string Message { get; }

    internal abstract bool IsValid(object? value);
}

internal sealed class AllowedValuesValidationRule : DeclarativeValidationRule
{
    internal AllowedValuesValidationRule(IEnumerable<string> values)
    {
        Values = values.ToArray();
    }

    internal IReadOnlyList<string> Values { get; }

    internal override string Message => $"Value must be one of: {string.Join(", ", Values)}.";

    internal override bool IsValid(object? value) => value is null || (value is string text && Values.Contains(text, StringComparer.Ordinal));
}

internal sealed class NumericRangeValidationRule : DeclarativeValidationRule
{
    internal NumericRangeValidationRule(object minimum, object maximum)
    {
        Minimum = minimum;
        Maximum = maximum;
    }

    internal object Minimum { get; }

    internal object Maximum { get; }

    internal override string Message => $"Value must be between {Number(Minimum)} and {Number(Maximum)}.";

    internal override bool IsValid(object? value)
    {
        if (value is null)
        {
            return true;
        }

        var comparable = (IComparable)value;
        return comparable.CompareTo(Minimum) >= 0 && comparable.CompareTo(Maximum) <= 0;
    }

    private static string Number(object value) => Convert.ToString(value, CultureInfo.InvariantCulture)
        ?? throw new InvalidOperationException("A numeric range bound cannot be null.");
}

internal sealed class DateRangeValidationRule : DeclarativeValidationRule
{
    internal DateRangeValidationRule(DateTime minimum, DateTime maximum)
    {
        Minimum = minimum;
        Maximum = maximum;
    }

    internal DateTime Minimum { get; }

    internal DateTime Maximum { get; }

    internal override string Message => $"Date must be between {Date(Minimum)} and {Date(Maximum)}.";

    internal override bool IsValid(object? value) => value is null || (value is DateTime date && date >= Minimum && date <= Maximum);

    private static string Date(DateTime value) => value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
}
