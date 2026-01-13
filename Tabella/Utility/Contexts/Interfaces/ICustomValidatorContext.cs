using Tabella.Utility.Providers.Interfaces;

namespace Tabella.Utility.Contexts.Interfaces;

/// <summary>
/// Interface for custom validator context.
/// </summary>
public interface ICustomValidatorContext
{
    /// <summary>
    /// Gets the templates provider for Tabella messages.
    /// </summary>
    ITabellaMessageTemplatesProvider Templates { get; }
}