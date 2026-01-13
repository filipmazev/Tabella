using Tabella.Utility.Factories.Interfaces;
using Tabella.Utility.Providers.Interfaces;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Extensions.Options;
using Tabella.Services.Interfaces;
using MessageKit.Utility.Builders;
using Tabella.Utility.Contexts;
using DocumentFormat.OpenXml;
using Tabella.Configuration;
using Tabella.Data.Core;
using Tabella._Common;

namespace Tabella.Services;

/// <summary>
/// Service for processing tabular files such as Excel spreadsheets.
/// </summary>
/// <param name="options"></param>
/// <param name="templates"></param>
/// <param name="messageFactory"></param>
public partial class TabularFileProcessingService(
    IOptions<TabularFileImportMessageOptions> options,
    ITabellaMessageTemplatesProvider templates,
    ITabularFileImportMessageFactory messageFactory) 
    : ITabularFileProcessingService
{
    #region Import
    
    #region Processing Methods

    /// <summary>
    /// Processes a sheeted tabular file from the provided stream.
    /// </summary>
    /// <param name="stream"></param>
    /// <param name="sheetMappingsOrdered"></param>
    /// <param name="headerRowIndex"></param>
    /// <returns></returns>
    public TabularFileImportResult? ProcessSheetedFile(
        Stream stream,
        IEnumerable<KeyValuePair<string, TabularFileSheetMapping>> sheetMappingsOrdered,
        int headerRowIndex = 0)
    {
        using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(stream, false);
        WorkbookPart? workbookPart = spreadsheet.WorkbookPart;
        
        using IDisposable customCastContext = CustomCastContextAccessor.Use(new CustomCastContext(templates));
        using IDisposable customValidatorContext = CustomValidatorContextAccessor.Use(new CustomValidatorContext(templates));

        if(workbookPart is null) return null;

        TabularFileImportResult result = ProcessSheets(
            workbookPart: workbookPart,
            sheetMappingsOrdered: sheetMappingsOrdered,
            headerRowIndex: headerRowIndex);

        return result;
    }
    
    #endregion

    #region Post Processing Methods

    #region Generic Helper Methods

    /// <summary>
    /// Retrieves a post-processing result of the specified type from the provided dictionary.
    /// </summary>
    /// <param name="type"></param>
    /// <param name="postProcessingResults"></param>
    /// <typeparam name="T"></typeparam>
    /// <returns></returns>
    public T? GetFileImportPostProcessingResult<T>(Type type, ref Dictionary<Type, object?> postProcessingResults)
        where T : class
    {
        return postProcessingResults.TryGetValue(type, out object? value) ? value as T : null;
    }
    
    #endregion
    
    #region Generic Post Processing Validators

    /// <summary>
    /// Checks for an existing entry in the provided map based on the object's key.
    /// </summary>
    /// <param name="objectData"></param>
    /// <param name="existingEntitiesMap"></param>
    /// <param name="existingEntry"></param>
    /// <typeparam name="TAppItem"></typeparam>
    /// <typeparam name="TImportModel"></typeparam>
    /// <returns></returns>
    public bool ExistingEntryCheck<TAppItem, TImportModel>(
        TImportModel objectData,
        List<(TAppItem appItem, TImportModel importedModel)> existingEntitiesMap,
        out (TAppItem appItem, TImportModel appItemMappedToImportModel) existingEntry)
        where TAppItem : class
        where TImportModel : ImportModel
    {
        existingEntry = existingEntitiesMap.FirstOrDefault(entry => entry.importedModel.Key.HasValue 
                                                                    && entry.importedModel.Key.Value.Equals(objectData.Key));

        return existingEntry.appItem != null;
    }

    #endregion

    #region Tabular File Import Post Processing Validators

    /// <summary>
    /// Validates duplicate entries in the imported tabular file data.
    /// </summary>
    /// <param name="data"></param>
    /// <param name="messages"></param>
    /// <returns></returns>
    public bool ValidateDuplicatesOnTabularFileImport(
        List<TabularFileProcessedObject> data, 
        ref HashSet<TabularFileImportMessage> messages)
    {
        HashSet<UInt128> seenIds = [];

        bool hasDuplicates = false;

        foreach(TabularFileProcessedObject obj in data)
        {
            if (obj.Object is not ImportModel { Key: not null } model ||
                seenIds.Add(model.Key.Value)) continue;
            
            messages.Add(messageFactory.CreateMessage(
                messageBuilder: new MessageBuilder(templates.DuplicateEntry),
                sheetName: obj.SheetName,
                rowIndex: obj.RowIndex
            ));

            hasDuplicates = true;
        }

        return hasDuplicates;
    }

    /// <summary>
    /// Validates a date range on a tabular file import.
    /// </summary>
    /// <param name="targetEntityProperty"></param>
    /// <param name="dateRange"></param>
    /// <param name="processedObject"></param>
    /// <param name="messages"></param>
    /// <returns></returns>
    public bool ValidateDateRangeOnTabularFileImport(
        string targetEntityProperty,
        (DateTime startDate, DateTime? endDate) dateRange,
        TabularFileProcessedObject processedObject,
        ref HashSet<TabularFileImportMessage> messages)
    {
        if(IsValidDateRange(dateRange)) return true;

        processedObject.ColumnData.TryGetValue(targetEntityProperty, out TabularFileColumnData? objectPropertyInfo);

        messages.Add(messageFactory.CreateMessage(
            messageBuilder: new MessageBuilder(templates.InvalidDateRange)
                .With(TabellaConstants.MessagePlaceholderStartDate, dateRange.startDate.ToString(TabellaConstants.DateFormat))
                .With(TabellaConstants.MessagePlaceholderEndDate, dateRange.endDate?.ToString(TabellaConstants.DateFormat) ?? TabellaConstants.MessageRangeNoEndSymbol),
            sheetName: processedObject.SheetName,
            rowIndex: processedObject.RowIndex,
            cellData: objectPropertyInfo?.CellData
        ));

        return false;
    }

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
    public List<TImportModel> ValidateAndRetrieveItemsInRangeOnTabularFileImport<TAppItem, TImportModel>(
        UInt128 lookupKey,
        (DateTime startDate, DateTime? endDate) itemRange,
        string itemPropertyName,
        ref TabularFileImportPostProcessingResult<TAppItem, TImportModel> existingItems,
        Func<TImportModel, (DateTime startDate, DateTime? endDate)> dateRangeSelector,
        ref TabularFileProcessedObject processedObject,
        ref HashSet<TabularFileImportMessage> messages)
        where TAppItem : class
        where TImportModel : ImportModel
    {
        bool ignoreRange = itemRange.startDate.Equals(DateTime.MinValue) || itemRange.endDate.Equals(DateTime.MinValue);

        List<TImportModel> existingItemsForIdentifier = existingItems.ImportedModels
            .Where(entry => entry.LookupKey.HasValue 
            && entry.LookupKey.Equals(lookupKey)
            && (ignoreRange || IsNewDateRangeContainedInExistingDateRange(itemRange, dateRangeSelector(entry))))
            .ToList();

        bool itemFound = existingItemsForIdentifier.Any();

        if(itemFound) return existingItemsForIdentifier;

        processedObject.ColumnData.TryGetValue(itemPropertyName, out TabularFileColumnData? cellData);

        messages.Add(messageFactory.CreateMessage(
            messageBuilder: new MessageBuilder(templates.NoAccompanyingItemFound)
                .With(TabellaConstants.MessagePlaceholderCurrentValue, itemPropertyName),
            sheetName: processedObject.SheetName,
            rowIndex: processedObject.RowIndex,
            cellData: cellData?.CellData
        ));

        return existingItemsForIdentifier;
    }

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
    public bool ValidateOverlappingDateRangesOnTabularFileImport<TModel>(
        List<TabularFileProcessedObject> data,
        ref HashSet<TabularFileImportMessage> messages,
        Func<TModel, string>? groupKeySelector,
        Func<TModel, (DateTime startDate, DateTime? endDate)> dateRangeSelector,
        string startDatePropertyName,
        bool ignoreExactMatch = true)
        where TModel : class
    {
        Dictionary<string, List<(TabularFileProcessedObject processedObject, (DateTime startDate, DateTime? endDate) dateRange)>> dateRangeGroups;

        if(groupKeySelector is not null)
        {
            dateRangeGroups = data
                .Where(d => d.Object is TModel)
                .Select(d => new { processedObject = d, Entity = (TModel)d.Object! })
                .GroupBy(d => groupKeySelector(d.Entity))
                .ToDictionary(
                    group => group.Key,
                    group => group.Select(entry =>
                        (entry.processedObject, dateRange: dateRangeSelector(entry.Entity)))
                        .OrderBy(d => d.dateRange.startDate)
                        .ToList()
                );
        }
        else
        {
            dateRangeGroups = new Dictionary<string, List<(TabularFileProcessedObject processedObject, (DateTime startDate, DateTime? endDate) dateRange)>>()
            {
                [TabellaConstants.SheetedFileImportedEntityValidatorSingleGroupingKey] = data
                    .Where(d => d.Object is TModel)
                    .Select(d => (processedObject: d, dateRange: dateRangeSelector((TModel)d.Object!)))
                    .OrderBy(d => d.dateRange.startDate)
                    .ToList()
            };
        }

        bool overlapsExist = false;

        foreach(KeyValuePair<string, List<(TabularFileProcessedObject processedObject, (DateTime startDate, DateTime? endDate) dateRange)>> dateRangeGroup in dateRangeGroups)
        {
            if(dateRangeGroup.Value.Count == 1) continue;

            List<(DateTime startDate, DateTime? endDate)> existingDateRanges = [];

            foreach((TabularFileProcessedObject processedObject, (DateTime startDate, DateTime? endDate) dateRange) in dateRangeGroup.Value)
            {
                if(existingDateRanges.Any() && IsNewDateRangeOverlappingExistingDateRanges(dateRange, existingDateRanges, ignoreExactMatch: ignoreExactMatch))
                {
                    processedObject.ColumnData.TryGetValue(startDatePropertyName, out TabularFileColumnData? cellData);

                    messages.Add(messageFactory.CreateMessage(
                        messageBuilder: new MessageBuilder(templates.DateRangeOverlapsExisting)
                            .With(TabellaConstants.MessagePlaceholderStartDate, dateRange.startDate.ToString(TabellaConstants.DateFormat))
                            .With(TabellaConstants.MessagePlaceholderEndDate, dateRange.endDate?.ToString(TabellaConstants.DateFormat) ?? TabellaConstants.MessageRangeNoEndSymbol),
                        sheetName: processedObject.SheetName,
                        rowIndex: processedObject.RowIndex,
                        cellData: cellData?.CellData
                    ));

                    overlapsExist = true;
                }

                existingDateRanges.Add(dateRange);
            }
        }

        return overlapsExist;
    }

    #endregion

    #endregion
    
    #endregion
    
    #region Export
    
    /// <summary>
    /// Exports sheets to a tabular file format.
    /// </summary>
    /// <param name="outputStream"></param>
    /// <param name="sheetMappingsOrdered"></param>
    /// <param name="fetchEntitiesAsync"></param>
    /// <param name="cancellationToken"></param>
    /// <returns></returns>
    public async Task<TabularFileExportResult> ExportSheets(
        Stream outputStream,
        IEnumerable<KeyValuePair<string, TabularFileSheetMapping>> sheetMappingsOrdered,
        Func<Type, CancellationToken, Task<IEnumerable<object>>> fetchEntitiesAsync,
        CancellationToken cancellationToken = default)
    {
        TabularFileExportResult result = new()
        {
            IsSuccess = true,
            Warnings = []
        };

        try
        {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Create(outputStream, SpreadsheetDocumentType.Workbook, true);
            WorkbookPart workbookPart = spreadsheet.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = CreateStylesheet();
            stylesPart.Stylesheet.Save();

            Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

            uint sheetId = 1;

            foreach ((string sheetNameRaw, TabularFileSheetMapping mapping) in sheetMappingsOrdered)
            {
                await ProcessSheets(workbookPart, sheets, sheetNameRaw, mapping, sheetId++, fetchEntitiesAsync, result, cancellationToken);
            }

            workbookPart.Workbook.Save();
            return result;
        }
        catch
        {
            result.IsSuccess = false;
            return result;
        }
    }
    
    #endregion
}