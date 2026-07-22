namespace CellSharp.Internal;

internal sealed class ValidationRule
{
    internal ValidationRule(Func<object?, bool> predicate, string message)
    {
        Predicate = predicate;
        Message = message;
    }

    internal Func<object?, bool> Predicate { get; }

    internal string Message { get; }
}
