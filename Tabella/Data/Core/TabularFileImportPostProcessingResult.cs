using Tabella.Enums;

namespace Tabella.Data.Core;

/// <summary>
/// Represents the result of post-processing after importing tabular file data.
/// </summary>
/// <typeparam name="TAppItem"></typeparam>
/// <typeparam name="TImportModel"></typeparam>
public class TabularFileImportPostProcessingResult<TAppItem, TImportModel>
    where TAppItem : class
    where TImportModel : ImportModel
{
    /// <summary>
    /// The list of existing database entities matched with their corresponding imported models.
    /// </summary>
    public List<(TAppItem appItem, TImportModel importedModel)> ExistingDbEntities { get; set; } = [];
    
    /// <summary>
    /// The list of entities to be created, each paired with its corresponding imported model.
    /// </summary>
    public List<(TAppItem appItem, TImportModel importedModel)> CreatingEntities { get; set; } = [];
    
    /// <summary>
    /// The list of entities to be updated, each paired with its corresponding imported model.
    /// </summary>
    public List<(TAppItem appItem, TImportModel importedModel)> UpdatingEntities { get; set; } = [];

    /// <summary>
    /// The list of all imported models.
    /// </summary>
    public List<TImportModel> ImportedModels { get; set; } = [];
    
    /// <summary>
    /// The list of pending import models along with their associated application items, processed objects, and action types.
    /// </summary>
    public List<(TAppItem? appItem, TImportModel importedModel, TabularFileProcessedObject processedObject, ActionTypeEnum action)> PendingImportModels { get; set; } = [];

    /// <summary>
    /// The combined list of entities after merging existing, updating, and creating entities.
    /// </summary>
    public List<(TAppItem appItem, TImportModel importedModel)> CombinedEntities { get; private set; } = [];

    /// <summary>
    /// Generates the combined list of entities by merging existing, updating, and creating entities.
    /// </summary>
    public void GenerateCombinedEntities()
    {
        CombinedEntities =
            ExistingDbEntities
                .Concat(UpdatingEntities)
                .GroupBy(item => item.GetHashCode()) 
                .Select(group => group.Last())         
                .Concat(CreatingEntities)
                .ToList();
    }
}