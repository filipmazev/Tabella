using Tabella.Data.Core;

namespace Tabella.Services.Interfaces;

/// <summary>
/// Service interface for processing tabular files (e.g., Excel, CSV).
/// </summary>
public interface ITabularFileProcessingService
{
    #region Import
    
    #region File Processing Methods

    /// <summary>
    /// Processes a sheeted tabular file (e.g., Excel) from the provided stream using the specified sheet mappings.
    /// </summary>
    /// <param name="stream"></param>
    /// <param name="sheetMappingsOrdered"></param>
    /// <param name="headerRowIndex"></param>
    /// <returns></returns>
    TabularFileImportResult? ProcessSheetedFile(
        Stream stream,
        IEnumerable<KeyValuePair<string, TabularFileSheetMapping>> sheetMappingsOrdered,
        int headerRowIndex = 0);
    
    #endregion
    
    #region Post Processing Methods

    #region Generic Helper Methods

    /// <summary>
    /// Retrieves the post-processing result of a specific type from the provided dictionary.
    /// </summary>
    /// <param name="type"></param>
    /// <param name="postProcessingResults"></param>
    /// <typeparam name="T"></typeparam>
    /// <returns></returns>
    T? GetFileImportPostProcessingResult<T>(Type type, ref Dictionary<Type, object?> postProcessingResults)
        where T : class;
    
    #endregion
    
    #region Generic Post Processing Validators

    /// <summary>
    /// Checks for an existing entry in the provided existing entities map that matches the given object data.
    /// </summary>
    /// <param name="objectData"></param>
    /// <param name="existingEntitiesMap"></param>
    /// <param name="existingEntry"></param>
    /// <typeparam name="TAppItem"></typeparam>
    /// <typeparam name="TImportModel"></typeparam>
    /// <returns></returns>
    bool ExistingEntryCheck<TAppItem, TImportModel>(
        TImportModel objectData,
        List<(TAppItem appItem, TImportModel importedModel)> existingEntitiesMap,
        out (TAppItem appItem, TImportModel appItemMappedToImportModel) existingEntry)
        where TAppItem : class
        where TImportModel : ImportModel;

    #endregion

    #region Tabular File Import Post Processing Validators

    /// <summary>
    /// Validates duplicate entries in the provided tabular file processed objects.
    /// </summary>
    /// <param name="data"></param>
    /// <param name="messages"></param>
    /// <returns></returns>
    bool ValidateDuplicatesOnTabularFileImport(
        List<TabularFileProcessedObject> data,
        ref HashSet<TabularFileImportMessage> messages);

    /// <summary>
    /// Validates a date range on a tabular file import.
    /// </summary>
    /// <param name="targetEntityProperty"></param>
    /// <param name="dateRange"></param>
    /// <param name="processedObject"></param>
    /// <param name="messages"></param>
    /// <returns></returns>
    bool ValidateDateRangeOnTabularFileImport(
        string targetEntityProperty,
        (DateTime startDate, DateTime? endDate) dateRange,
        TabularFileProcessedObject processedObject,
        ref HashSet<TabularFileImportMessage> messages);

    /// <summary>
    /// Validates and retrieves items within a specified date range on a tabular file import.
    /// </summary>
    /// <param name="lookupKey"></param>
    /// <param name="itemRange"></param>
    /// <param name="itemPropertyName"></param>
    /// <param name="existingItems"></param>
    /// <param name="dateRangeSelector"></param>
    /// <param name="processedObject"></param>
    /// <param name="messages"></param>
    /// <typeparam name="TAppItem"></typeparam>
    /// <typeparam name="TImportModel"></typeparam>
    /// <returns></returns>
    List<TImportModel> ValidateAndRetrieveItemsInRangeOnTabularFileImport<TAppItem, TImportModel>(
        UInt128 lookupKey,
        (DateTime startDate, DateTime? endDate) itemRange,
        string itemPropertyName,
        ref TabularFileImportPostProcessingResult<TAppItem, TImportModel> existingItems,
        Func<TImportModel, (DateTime startDate, DateTime? endDate)> dateRangeSelector,
        ref TabularFileProcessedObject processedObject,
        ref HashSet<TabularFileImportMessage> messages)
        where TAppItem : class
        where TImportModel : ImportModel;

    /// <summary>
    /// Validates overlapping date ranges in the provided tabular file processed objects.
    /// </summary>
    /// <param name="data"></param>
    /// <param name="messages"></param>
    /// <param name="groupKeySelector"></param>
    /// <param name="dateRangeSelector"></param>
    /// <param name="startDatePropertyName"></param>
    /// <param name="ignoreExactMatch"></param>
    /// <typeparam name="TModel"></typeparam>
    /// <returns></returns>
    bool ValidateOverlappingDateRangesOnTabularFileImport<TModel>(
        List<TabularFileProcessedObject> data,
        ref HashSet<TabularFileImportMessage> messages,
        Func<TModel, string>? groupKeySelector,
        Func<TModel, (DateTime startDate, DateTime? endDate)> dateRangeSelector,
        string startDatePropertyName,
        bool ignoreExactMatch = true)
        where TModel : class;

    #endregion

    #endregion
    
    #endregion
    
    #region Export
    
    /// <summary>
    /// Exports data to a tabular file with multiple sheets based on the provided sheet mappings and data fetching function.
    /// </summary>
    /// <param name="outputStream"></param>
    /// <param name="sheetMappingsOrdered"></param>
    /// <param name="fetchEntitiesAsync"></param>
    /// <param name="cancellationToken"></param>
    /// <returns></returns>
    Task<TabularFileExportResult> ExportSheets(
        Stream outputStream,
        IEnumerable<KeyValuePair<string, TabularFileSheetMapping>> sheetMappingsOrdered,
        Func<Type, CancellationToken, Task<IEnumerable<object>>> fetchEntitiesAsync,
        CancellationToken cancellationToken = default);
    
    #endregion
}