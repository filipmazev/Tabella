using Tabella.Utility.Contexts.Interfaces;

namespace Tabella.Utility.Contexts;

/// <summary>
/// Accessor for the current CustomCastContext.
/// </summary>
public static class CustomCastContextAccessor
{
    // ReSharper disable once InconsistentNaming
    private static readonly AsyncLocal<ICustomCastContext?> _current = new();

    /// <summary>
    /// Gets the current CustomCastContext.
    /// </summary>
    /// <exception cref="InvalidOperationException"></exception>
    public static ICustomCastContext Current =>
        _current.Value ?? throw new InvalidOperationException(
            "CustomCastContext is not initialized.");

    /// <summary>
    /// Sets the current CustomCastContext for the scope of the returned IDisposable.
    /// </summary>
    /// <param name="context"></param>
    /// <returns></returns>
    public static IDisposable Use(ICustomCastContext context)
    {
        ICustomCastContext? previous = _current.Value;
        _current.Value = context;
        return new Scope(() => _current.Value = previous);
    }

    private sealed class Scope(Action onDispose) : IDisposable
    {
        public void Dispose() => onDispose();
    }
}