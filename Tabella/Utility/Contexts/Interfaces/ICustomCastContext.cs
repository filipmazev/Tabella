using Tabella.Utility.Providers.Interfaces;

namespace Tabella.Utility.Contexts.Interfaces;

/// <summary>
/// Interface for custom cast context.
/// </summary>
public interface ICustomCastContext
{
    /// <summary>
    /// Gets the templates provider for Tabella messages.
    /// </summary>
    ITabellaMessageTemplatesProvider Templates { get; }
}