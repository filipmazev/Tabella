using Tabella.Utility.Providers.Interfaces;
using Tabella.Utility.Contexts.Interfaces;

namespace Tabella.Utility.Contexts;

internal sealed class CustomValidatorContext(ITabellaMessageTemplatesProvider templates) : ICustomValidatorContext
{
    public ITabellaMessageTemplatesProvider Templates { get; } = templates;
}