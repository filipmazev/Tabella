using Tabella.Utility.Providers.Interfaces;
using Tabella.Utility.Contexts.Interfaces;

namespace Tabella.Utility.Contexts;

internal sealed class CustomCastContext(ITabellaMessageTemplatesProvider templates) : ICustomCastContext
{
    public ITabellaMessageTemplatesProvider Templates { get; } = templates;
}