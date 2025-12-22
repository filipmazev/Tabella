using Microsoft.Extensions.DependencyInjection;
using Tabella.Utility.Providers.Interfaces;
using Tabella.Utility.Factories.Interfaces;
using Tabella.Utility.Helpers.Interfaces;
using Tabella.Services.Interfaces;
using Tabella.Utility.Factories;
using Tabella.Utility.Providers;
using Tabella.Utility.Helpers;
using Tabella.Configuration;
using Tabella.Services;
using MessageKit;

namespace Tabella;

/// <summary>
/// Extension methods for registering tabular file import message services.
/// </summary>
public static class TabularFileImportMessageServiceCollectionExtensions
{
    /// <summary>
    /// Registers services for tabular file import messages.
    /// </summary>
    /// <param name="services"></param>
    /// <param name="configure"></param>
    /// <param name="configureMessageOptions"></param>
    /// <returns></returns>
    public static IServiceCollection AddTabularFileImportMessages(
        this IServiceCollection services,
        Action<TabularFileImportMessageOptions> configure,
        Action<MessageOptions> configureMessageOptions)
    {
        services.Configure(configure);
        
        services.AddSingleton<ITabellaMessageTemplatesProvider, TabellaMessageTemplatesProvider>();
        services.AddSingleton<ITabularFileImportMessageFactory, TabularFileImportMessageFactory>();
       
        services.AddTransient<ITabularFileProcessingService, TabularFileProcessingService>();
        services.AddTransient<ICustomCasts, CustomCasts>();
        services.AddTransient<ICustomValidators, CustomValidators>();
        
        services.AddMessages(configureMessageOptions.Invoke);
        
        return services;
    }
}