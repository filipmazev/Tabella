using Tabella.Utility.Contexts.Interfaces;

namespace Tabella.Utility.Contexts;

/// <summary>
/// Accessor for the current CustomValidatorContext.
/// </summary>
public static class CustomValidatorContextAccessor
{
    // ReSharper disable once InconsistentNaming
    private static readonly AsyncLocal<ICustomValidatorContext?> _current = new();

    /// <summary>
    /// Gets the current CustomValidatorContext.
    /// </summary>
    /// <exception cref="InvalidOperationException"></exception>
    public static ICustomValidatorContext Current =>
        _current.Value ?? throw new InvalidOperationException(
            "CustomValidatorContext is not initialized.");

    /// <summary>
    /// Sets the current CustomValidatorContext for the scope of the returned IDisposable.
    /// </summary>
    /// <param name="context"></param>
    /// <returns></returns>
    public static IDisposable Use(ICustomValidatorContext context)
    {
        ICustomValidatorContext? previous = _current.Value;
        _current.Value = context;
        return new Scope(() => _current.Value = previous);
    }

    private sealed class Scope(Action onDispose) : IDisposable
    {
        public void Dispose() => onDispose();
    }
}