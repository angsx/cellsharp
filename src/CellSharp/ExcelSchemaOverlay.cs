using System.Collections.ObjectModel;
using System.Linq.Expressions;
using System.Reflection;
using CellSharp.Internal;

namespace CellSharp;

/// <summary>Describes immutable per-operation changes to a typed schema.</summary>
public sealed class ExcelSchemaOverlay<T>
{
    internal ExcelSchemaOverlay(IReadOnlyDictionary<PropertyInfo, RuntimeColumnOverride> columns)
    {
        Columns = columns;
    }

    internal IReadOnlyDictionary<PropertyInfo, RuntimeColumnOverride> Columns { get; }
}

/// <summary>Builds an immutable per-operation overlay for a typed schema.</summary>
public sealed class ExcelSchemaOverlayBuilder<T>
{
    private readonly Dictionary<PropertyInfo, RuntimeColumnOverride> _columns = new();

    /// <summary>Overrides the effective header for this operation.</summary>
    public ExcelSchemaOverlayBuilder<T> Header<TValue>(Expression<Func<T, TValue>> property, string header)
    {
        if (string.IsNullOrWhiteSpace(header))
        {
            throw new ArgumentException("A runtime header is required.", nameof(header));
        }

        var propertyInfo = PropertyFrom(property);
        if (_columns.TryGetValue(propertyInfo, out var current) && current.Header is not null)
        {
            throw new InvalidOperationException($"A runtime header is already configured for property '{propertyInfo.Name}'.");
        }

        _columns[propertyInfo] = new RuntimeColumnOverride(header, current?.IsEnabled);
        return this;
    }

    /// <summary>Includes or excludes this column for this operation.</summary>
    public ExcelSchemaOverlayBuilder<T> Include<TValue>(Expression<Func<T, TValue>> property, bool enabled)
    {
        var propertyInfo = PropertyFrom(property);
        if (_columns.TryGetValue(propertyInfo, out var current) && current.IsEnabled is not null)
        {
            if (current.IsEnabled.Value != enabled)
            {
                throw new InvalidOperationException($"Runtime participation for property '{propertyInfo.Name}' is already configured.");
            }

            return this;
        }

        _columns[propertyInfo] = new RuntimeColumnOverride(current?.Header, enabled);
        return this;
    }

    /// <summary>Creates an immutable overlay that can be safely reused across operations.</summary>
    public ExcelSchemaOverlay<T> Build() => new(
        new ReadOnlyDictionary<PropertyInfo, RuntimeColumnOverride>(
            new Dictionary<PropertyInfo, RuntimeColumnOverride>(_columns)));

    private static PropertyInfo PropertyFrom<TValue>(Expression<Func<T, TValue>> expression)
    {
        if (expression is null)
        {
#if NET8_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(expression);
#else
            throw new ArgumentNullException(nameof(expression));
#endif
        }

        if (expression.Body is not MemberExpression { Member: PropertyInfo property, Expression: ParameterExpression })
        {
            throw new ArgumentException("A runtime schema column must select a direct property access.", nameof(expression));
        }

        return property;
    }
}
