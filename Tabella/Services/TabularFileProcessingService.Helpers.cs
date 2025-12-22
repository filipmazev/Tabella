using System.ComponentModel.DataAnnotations;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;
using MessageKit.Utility.Extensions;
using MessageKit.Utility.Builders;
using Tabella.Utility.Delegates;
using Tabella.Utility.Helpers;
using Tabella.Data.Interfaces;
using DocumentFormat.OpenXml;
using System.Globalization;
using Tabella.Data.Results;
using System.Reflection;
using Tabella.Data.Core;
using Tabella._Common;
using Tabella.Enums;

namespace Tabella.Services;

public partial class TabularFileProcessingService
{
    #region Import
    
    #region Sheet Processing

    private TabularFileImportResult ProcessSheets(
        WorkbookPart workbookPart,
        IEnumerable<KeyValuePair<string, TabularFileSheetMapping>> sheetMappingsOrdered,
        int headerRowIndex = 0)
    {
        TabularFileImportResult result = new()
        {
            ProcessedObjects = new Dictionary<Type, List<TabularFileProcessedObject>>(),
            Messages = []
        };

        Dictionary<string, TabularFileSheetMapping> sheetMappings = sheetMappingsOrdered.Select(item =>
            {
                item.Value.OriginalSheetName = item.Key;
                return item;
            })
            .ToDictionary(
                kvp => kvp.Key.Length > TabellaConstants.ExcelSheetNameMaxCharacterCount
                    ? CustomFormatters.ToCleanString(kvp.Key[..TabellaConstants.ExcelSheetNameMaxCharacterCount])
                    : CustomFormatters.ToCleanString(kvp.Key),
                kvp => kvp.Value
            );

        List<Sheet> sheets = workbookPart.Workbook.Descendants<Sheet>()
            .Where((sheet) => sheet.State is null || !sheet.State.HasValue
                || (sheet.State.Value != SheetStateValues.Hidden && sheet.State.Value != SheetStateValues.VeryHidden))
            .ToList();

        bool sheetFoundDuringProcessing = false;

        foreach(Sheet sheet in sheets)
        {
            if(sheet.Name is null || !sheet.Name.HasValue || sheet.Id is null || !sheet.Id.HasValue || sheet.Id.Value is null) continue;

            if(workbookPart.GetPartById(sheet.Id.Value) is not WorksheetPart worksheetPart) continue;

            string worksheetName = CustomFormatters.ToCleanString(sheet.Name.Value!);

            if(!sheetMappings.TryGetValue(worksheetName, out TabularFileSheetMapping? currentSheetMapping)) continue;

            currentSheetMapping.FoundInFile = true;
            sheetMappings[worksheetName] = currentSheetMapping;

            ValidateColumnMappings(currentSheetMapping.ColumnMappings);

            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData? sheetData = worksheet.GetFirstChild<SheetData>();

            if(sheetData is null) continue;

            sheetFoundDuringProcessing = true;

            Type entityType = currentSheetMapping.EntityType;

            ProcessRows(
                sheetData: sheetData,
                sheetName: sheet.Name.Value,
                workbookPart: workbookPart,
                columnMappings: currentSheetMapping.ColumnMappings.ToDictionary(kvp => CustomFormatters.ToCleanString(kvp.Key), kvp => kvp.Value),
                entityType: entityType,
                result: ref result,
                headerRowIndex: headerRowIndex);
        }

        foreach (TabularFileImportMessage message in from item in sheetMappings where !item.Value.FoundInFile select messageFactory.CreateMessage(
                     messageBuilder: 
                     new MessageBuilder( templates.WorksheetNotFound)
                         .With(nameof(TabularFileProcessedObject.SheetName), item.Value.OriginalSheetName),
                     sheetName: item.Key))
        {
            result.Messages.Add(message);
        }

        if (sheetFoundDuringProcessing) return result;
        {
            TabularFileImportMessage message = messageFactory.CreateMessage(messageBuilder: new MessageBuilder(templates.NoWorksheetsFound));
            result.Messages.Add(message);
        }

        return result;
    }

    #endregion

    #region Rows Processing

    private void ProcessRows(
        SheetData sheetData,
        string? sheetName,
        WorkbookPart workbookPart,
        Dictionary<string, TabularFileColumnMapping> columnMappings,
        Type entityType,
        ref TabularFileImportResult result,
        int headerRowIndex = 0)
    {
        if(!result.ProcessedObjects.TryGetValue(entityType, out List<TabularFileProcessedObject>? currentProcessedObjects))
        {
            currentProcessedObjects = [];
            result.ProcessedObjects[entityType] = currentProcessedObjects;
        }

        List<TabularFileProcessedObject> processedObjects = [];

        List<Row> rows = sheetData.Elements<Row>()
            .Where(row => row.HasChildren && (row.Hidden is null || !row.Hidden.Value))
            .ToList();

        IEnumerable<Column> columns = sheetData.Elements<Column>().Where(c => c.Hidden is not null && c.Hidden.Value);

        Dictionary<uint, bool> hiddenColumns = new();

        foreach(Column item in columns)
        {
            if(item.Min is null || item.Max is null) continue;

            for(uint i = item.Min.Value; i <= item.Max.Value; i++)
            {
                hiddenColumns[i] = true;
            }
        }

        Row? headerRow = rows.ElementAtOrDefault(headerRowIndex);

        if(headerRow is null)
        {
            TabularFileImportMessage message = messageFactory.CreateMessage(
                    messageBuilder: new MessageBuilder(templates.HeaderNotFoundAtIndex)
                        .With(TabellaConstants.MessagePlaceholderIndex, headerRowIndex.ToString()),
                    sheetName: sheetName);

            result.Messages.Add(message);

            return;
        }

        Dictionary<string, TabularFileImportHeaderCell> headerCells = ProcessHeaderCells(
            cells: CollectCellValues(workbookPart, hiddenColumns, headerRow, headerRowIndex, sheetName),
            columnMappings: ref columnMappings,
            result: result);

        for(int rowIndex = headerRowIndex + 1; rowIndex < rows.Count; rowIndex++)
        {
            TabularFileProcessedObject? cellProcessingResult = ProcessCells(
                entityType: entityType,
                cells: CollectCellValues(workbookPart, hiddenColumns, rows[rowIndex], rowIndex, sheetName),
                headerCells: headerCells,
                result: result,
                rowIndex: rowIndex);

            if(cellProcessingResult?.Object is not TabularFileImportModel model) continue;

            model.GenerateKeys(rowIndex);
            model.SetCellInfo(cellProcessingResult);

            processedObjects.Add(cellProcessingResult);
        }

        foreach (TabularFileImportMessage? message in from item 
                    in columnMappings 
                    where !item.Value.FoundInFile 
                    select messageFactory.CreateMessage(
                     messageBuilder: new MessageBuilder(templates.ColumnNotFound)
                         .With(nameof(TabularFileCellData.ColumnName), item.Value.OriginalHeaderName),
                     sheetName: sheetName))
        {
            result.Messages.Add(message);
        }

        currentProcessedObjects.AddRange(processedObjects);
    }

    #endregion

    #region Cells Processing

    private Dictionary<string, TabularFileCell> CollectCellValues(
        WorkbookPart workbookPart,
        Dictionary<uint, bool> hiddenColumns,
        Row row,
        int rowIndex,
        string? sheetName = null)
    {
        return row.Elements<Cell>()
            .Where(cell => cell.CellReference?.Value != null)
            .Select((cell, index) =>
            {
                if(hiddenColumns.ContainsKey((uint)(index + 1))) return default;

                string? cellRef = GetColumnFromCellReference(cell.CellReference?.Value);

                if(cellRef is null) return default;

                object? cellValue = GetCellValue(cell, workbookPart);

                return new KeyValuePair<string, TabularFileCell>(
                    cellRef,
                    new TabularFileCell
                    {
                        CellValue = cellValue is string str && string.IsNullOrEmpty(CustomFormatters.ToCleanString(str)) ? null : cellValue,
                        ColumnIndex = index + 1,
                        RowIndex = rowIndex,
                        SheetName = sheetName
                    }
                );

            })
            .Where(pair => !string.IsNullOrEmpty(pair.Key))
            .ToDictionary(pair => pair.Key, pair => pair.Value);
    }

    #region Sheeted File Cell Value Parsing

    private object? GetCellValue(Cell cell, WorkbookPart workbookPart)
    {
        string value = cell.CellValue?.Text ?? string.Empty;

        if(cell.DataType is null) 
            return TryParseNumericValue(cell, value, ref workbookPart);

        if(cell.DataType.Value.Equals(CellValues.Boolean)) 
            return GetCellsBooleanValue(value);

        if( cell.DataType.Value.Equals(CellValues.SharedString)) 
            return GetCellsSharedStringValue(value, ref workbookPart);
        
        return cell.DataType.Value.Equals(CellValues.InlineString)
            ? GetCellInlineStringValue(cell)
            : TryParseNumericValue(cell, value, ref workbookPart);
    }

    private static bool? GetCellsBooleanValue(string value)
    {
        if(bool.TryParse(value, out bool boolValue)) return boolValue;
        if(int.TryParse(value, out int intValue)) return intValue != 0;

        return null;
    }

    private static string? GetCellsSharedStringValue(string value, ref WorkbookPart workbookPart)
    {
        if(!int.TryParse(value, out int sharedStringId)) return null;

        SharedStringTablePart? sharedStringTablePart = workbookPart.SharedStringTablePart;

        return sharedStringTablePart?.SharedStringTable.Elements<SharedStringItem>()
            .ElementAt(sharedStringId).InnerText;
    }
    
    private string GetCellInlineStringValue(Cell cell)
    {
        InlineString? inlineString = cell.GetFirstChild<InlineString>();
        return inlineString?.Text?.Text ?? string.Empty;
    }
    
    private object TryParseNumericValue(Cell cell, string value, ref WorkbookPart workbookPart)
    {
        return double.TryParse(value, CultureInfo.InvariantCulture, out double numericValue)
            ? TryParseDateValue(cell, numericValue, ref workbookPart)
            : value;
    }

    private static object TryParseDateValue(Cell cell, double numericValue, ref WorkbookPart workbookPart)
    {
        if(cell.StyleIndex?.InnerText == null
           || workbookPart.WorkbookStylesPart?.Stylesheet.CellFormats is null) return numericValue;

        CellFormat? cellFormat = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ChildElements[int.Parse(cell.StyleIndex.InnerText)] as CellFormat;

        if(cellFormat?.NumberFormatId is null) return numericValue;

        uint formatId = cellFormat.NumberFormatId.Value;
        bool isDate = TabellaConstants.OpenXmlStandardDateFormatIds.Contains(formatId);

        if(!isDate)
        {
            NumberingFormats? numberingFormats = workbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats;
            NumberingFormat? numberingFormat = numberingFormats?.Elements<NumberingFormat>().FirstOrDefault(nf => nf.NumberFormatId == cellFormat.NumberFormatId);

            string? formatCode = numberingFormat?.FormatCode?.Value;

            if(formatCode != null && TabellaConstants.OpenXmlDateFormats.Contains(formatCode))
            {
                isDate = true;
            }
        }

        if(!isDate) return numericValue;

        DateTime baseDate = TabellaConstants.ExcelBaseDate;

        if(numericValue > TabellaConstants.ExcelLeapYearBug) numericValue--;

        return DateTime.SpecifyKind(baseDate.AddDays(numericValue - 1), DateTimeKind.Utc);
    }

    private static string? GetColumnFromCellReference(string? cellReference)
    {
        if(cellReference is null) return null;

        Match match = Regex.Match(cellReference, @"^[A-Za-z]+");
        return match.Success ? match.Value : null;
    }

    #endregion

    #endregion
    
    #region Tabular File Processing Helpers

    #region Tabular File Import Generic Cell Value Processing

    private void ProcessCellValue(
        object? cellValue,
        TabularFileColumnMapping columnMapping,
        object objectData,
        ref HashSet<TabularFileImportMessage> messages,
        ref TabularFileCellData cellData,
        string? sheetName,
        int? rowIndex,
        out (bool isValid, bool isRequired, bool messageLoggedForCell) cellValidationResult)
    {
        cellValidationResult.isValid = true;
        cellValidationResult.isRequired = false;
        cellValidationResult.messageLoggedForCell = false;

        Type objectPropertyType = columnMapping.ColumnData.ObjectPropertyType;

        string targetEntityProperty = columnMapping.ColumnData.TargetEntityProperty;

        PropertyInfo? propertyInfo = objectData.GetType().GetProperty(targetEntityProperty, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);

        if(propertyInfo is null || !propertyInfo.CanWrite || !objectPropertyType.IsAssignableFrom(propertyInfo.PropertyType)) return;

        RequiredAttribute? requiredAttribute = propertyInfo.GetCustomAttribute<RequiredAttribute>();

        cellValidationResult.isRequired = requiredAttribute is not null;

        if(cellValue is null)
        {
            if(!cellValidationResult.isRequired) return;

            cellValidationResult.messageLoggedForCell = true;

            messages.Add(messageFactory.CreateMessage(
                messageBuilder: new MessageBuilder(templates.RequiredField)
                    .With(nameof(TabularFileCellData.ColumnName), cellData.ColumnName ?? string.Empty),
                sheetName: sheetName,
                rowIndex: rowIndex,
                cellData: cellData));

            return;
        }

        object? currentValue = cellValue;

        if(columnMapping.CustomCast is not null)
        {
            object? castedValue = HandleCustomCastScenario(
                customCast: columnMapping.CustomCast,
                cellValue: cellValue,
                objectPropertyType: objectPropertyType,
                targetEntityProperty: targetEntityProperty,
                objectData: objectData,
                messages: ref messages,
                cellData: ref cellData,
                sheetName: sheetName,
                rowIndex: rowIndex,
                out (bool isValid, bool messageLoggedForCell) cellCustomCastResult);

            cellValidationResult.isValid = cellValidationResult.isValid && cellCustomCastResult is { isValid: true, messageLoggedForCell: false };
            cellValidationResult.messageLoggedForCell = cellValidationResult.messageLoggedForCell || cellCustomCastResult.messageLoggedForCell;

            currentValue = castedValue;
        }
        else
        {
            bool typeCheck = columnMapping.CellTypes.Contains(cellValue.GetType()) && objectPropertyType.IsInstanceOfType(cellValue);

            if(!typeCheck)
            {
                cellValidationResult.isValid = false;
            }
            else
            {
                HandleDefaultAssignmentScenario(
                    cellValue: cellValue,
                    columnMapping: columnMapping,
                    objectPropertyType: objectPropertyType,
                    targetEntityProperty: targetEntityProperty,
                    objectData: objectData,
                    out bool isDefaultAssignmentSuccessful);

                cellValidationResult.isValid = cellValidationResult.isValid && isDefaultAssignmentSuccessful;
            }
        }

        if(!cellValidationResult.isValid) return;

        if(columnMapping.CustomValidators.Any())
        {
             HandleCustomValidators(
                customValidators: columnMapping.CustomValidators,
                cellValue: currentValue,
                messages: ref messages,
                cellData: ref cellData,
                sheetName: sheetName,
                rowIndex: rowIndex,
                out (bool isValid, bool messageLoggedForCell) cellCustomValidationResult);

            cellValidationResult.isValid = cellValidationResult.isValid && cellCustomValidationResult is { isValid: true, messageLoggedForCell: false };
            cellValidationResult.messageLoggedForCell = cellValidationResult.messageLoggedForCell || cellCustomValidationResult.messageLoggedForCell;
        }
    }

    private TabularFileProcessedObject? ProcessCells(
        Type entityType,
        Dictionary<string, TabularFileCell> cells,
        Dictionary<string, TabularFileImportHeaderCell> headerCells,
        TabularFileImportResult result,
        int rowIndex)
    {
        TabularFileProcessedObject processedObject = new()
        {
            Object = Activator.CreateInstance(entityType),
            SheetName = null,
            RowIndex = rowIndex + 1,
            ColumnData = new Dictionary<string, TabularFileColumnData>()
        };

        if(processedObject.Object is null) return null;

        bool rowHasAnyValues = false;
        bool rowHasAnyRequiredFields = false;

        HashSet<TabularFileImportMessage> messages = [];

        foreach(KeyValuePair<string, TabularFileImportHeaderCell> headerCellItem in headerCells)
        {
            object? cellValue = cells.TryGetValue(headerCellItem.Key, out TabularFileCell? currentCell) ? currentCell.CellValue : null;

            if(!rowHasAnyValues) rowHasAnyValues = cellValue is not null;

            TabularFileCellData cellData = new()
            {
                ColumnName = headerCellItem.Value.CellValue?.ToString(),
                ColumnIndex = headerCellItem.Value.ColumnIndex,
                CellValue = cellValue?.ToString()
            };

            TabularFileColumnData objectPropertiesInfo = new()
            {
                ObjectPropertyType = headerCellItem.Value.ColumnMapping.ColumnData.ObjectPropertyType,
                TargetEntityProperty = headerCellItem.Value.ColumnMapping.ColumnData.TargetEntityProperty,
                CellData = cellData
            };

            processedObject.SheetName = headerCellItem.Value.SheetName;

            processedObject.ColumnData.Add(headerCellItem.Value.ColumnMapping.ColumnData.TargetEntityProperty, objectPropertiesInfo);

            ProcessCellValue(
                cellValue: cellValue,
                columnMapping: headerCellItem.Value.ColumnMapping,
                objectData: processedObject.Object,
                messages: ref messages,
                cellData: ref cellData,
                sheetName: headerCellItem.Value.SheetName,
                rowIndex: processedObject.RowIndex,
                out (bool isValid, bool isRequired, bool messageLoggedForCell) cellValidationResult);

            if(!rowHasAnyRequiredFields) rowHasAnyRequiredFields = cellValidationResult.isRequired;

            if(cellValidationResult.isValid || cellValidationResult.messageLoggedForCell) continue;

            TabularFileImportMessage message = messageFactory.CreateMessage(
                messageBuilder: new MessageBuilder(templates.CellProcessingError),
                sheetName: processedObject.SheetName,
                rowIndex: processedObject.RowIndex,
                cellData: cellData);

            messages.Add(message);
        }

        if(rowHasAnyValues && messages.Any()) result.Messages.UnionWith(messages);

        return (!messages.Any(item => item.InformationType.IsNegative()))
                && processedObject.ColumnData.Any()
                && (rowHasAnyValues || !rowHasAnyRequiredFields)
            ? processedObject
            : null;
    }

    private Dictionary<string, TabularFileImportHeaderCell> ProcessHeaderCells(
        Dictionary<string, TabularFileCell> cells,
        ref Dictionary<string, TabularFileColumnMapping> columnMappings,
        TabularFileImportResult result)
    {
        Dictionary<string, TabularFileImportHeaderCell> headerCells = new();

        foreach(KeyValuePair<string, TabularFileCell> item in cells)
        {
            if(item.Value.CellValue is null) continue;

            string? originalValue = item.Value.CellValue.ToString();
            string formattedValue = CustomFormatters.ToCleanString(item.Value.CellValue);

            if(string.IsNullOrEmpty(originalValue) || string.IsNullOrEmpty(formattedValue)) continue;

            if(columnMappings.TryGetValue(formattedValue, out TabularFileColumnMapping? columnMapping))
            {
                columnMapping.FoundInFile = true;
                columnMapping.OriginalHeaderName = originalValue;
                columnMappings[formattedValue] = columnMapping;

                TabularFileImportHeaderCell headerCell = new()
                {
                    CellValue = originalValue,
                    ColumnIndex = item.Value.ColumnIndex,
                    RowIndex = item.Value.RowIndex,
                    SheetName = item.Value.SheetName,
                    FormatedValue = formattedValue,
                    ColumnMapping = columnMapping
                };

                headerCells.TryAdd(item.Key, headerCell);
            }
            else
            {
                TabularFileImportMessage message = messageFactory.CreateMessage(
                    messageBuilder: new MessageBuilder(templates.WrongColumnInFile)
                        .With(nameof(TabularFileCellData.ColumnName), originalValue),
                    sheetName: item.Value.SheetName,
                    rowIndex: item.Value.RowIndex + 1,
                    cellData: new TabularFileCellData()
                    {
                        ColumnName = originalValue,
                        ColumnIndex = item.Value.ColumnIndex + 1,
                        CellValue = originalValue
                    });

                result.Messages.Add(message);
            }
        }

        return headerCells;
    }

    #region Cell Value Processing Handlers

    private object? HandleCustomCastScenario(
        CustomCastDelegate customCast,
        object cellValue,
        Type objectPropertyType,
        string targetEntityProperty,
        object objectData,
        ref HashSet<TabularFileImportMessage> messages,
        ref TabularFileCellData cellData,
        string? sheetName,
        int? rowIndex,
        out (bool isValid, bool messageLoggedForCell) cellCustomCastResult)
    {
        CustomCastResult customCastResult = customCast(cellValue);

        cellCustomCastResult.isValid = customCastResult.IsValid;
        cellCustomCastResult.messageLoggedForCell = false;

        if(customCastResult.CastedValue is not null && objectPropertyType.IsInstanceOfType(customCastResult.CastedValue))
        {
            PropertyInfo? propertyInfo = objectData.GetType().GetProperty(targetEntityProperty, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);

            if(propertyInfo is null || !propertyInfo.CanWrite) return customCastResult.CastedValue;

            try
            {
                propertyInfo.SetValue(objectData, customCastResult.CastedValue);
            }
            catch
            {
                cellCustomCastResult.isValid = false;
                return null;
            }
        }
        else
        if(!customCastResult.IsValid || customCastResult.CastedValue is MessageBuilder)
        {
            if(customCastResult.Message is null) return null;

            messages.Add(messageFactory.CreateMessage(
                messageBuilder: customCastResult.Message,
                sheetName: sheetName,
                rowIndex: rowIndex,
                cellData: cellData));

            cellCustomCastResult.messageLoggedForCell = true;

            return null;
        }

        return customCastResult.CastedValue;
    }

    private void HandleCustomValidators(
        List<CustomValidatorDelegate> customValidators,
        object? cellValue,
        ref HashSet<TabularFileImportMessage> messages,
        ref TabularFileCellData cellData,
        string? sheetName,
        int? rowIndex,
        out (bool isValid, bool messageLoggedForCell) cellCustomValidationResult)
    {
        cellCustomValidationResult.isValid = true;
        cellCustomValidationResult.messageLoggedForCell = false;

        foreach(CustomValidationResult validationResult in customValidators.Select(customValidator => customValidator(cellValue)))
        {
            if(validationResult.Message is not null)
            {
                messages.Add(messageFactory.CreateMessage(
                    messageBuilder: validationResult.Message,
                    sheetName: sheetName,
                    rowIndex: rowIndex,
                    cellData: cellData));

                cellCustomValidationResult.messageLoggedForCell = true;
            }

            if(validationResult.IsValid) continue;

            cellCustomValidationResult.isValid = false;
            return;
        }
    }

    private void HandleDefaultAssignmentScenario(
        object cellValue,
        TabularFileColumnMapping columnMapping,
        Type objectPropertyType,
        string targetEntityProperty,
        object objectData,
        out bool isDefaultAssignmentSuccessful)
    {
        isDefaultAssignmentSuccessful = true;

        if(!columnMapping.CellTypes.Contains(cellValue.GetType()) ||
           !objectPropertyType.IsInstanceOfType(cellValue)) return;

        PropertyInfo? propertyInfo = objectData.GetType().GetProperty(targetEntityProperty, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);

        if(propertyInfo is null || !propertyInfo.CanWrite) return;

        try
        {
            propertyInfo.SetValue(objectData, cellValue is string stringValue ? stringValue.Trim() : cellValue);
        }
        catch
        {
            isDefaultAssignmentSuccessful = false;
        }
    }

    #endregion

    #endregion

    #region Tabular File Import Generic Validators

    private void ValidateColumnMappings(Dictionary<string, TabularFileColumnMapping> columnMapping)
    {
        List<string> duplicateProperties = columnMapping.Values
            .GroupBy(mapping => mapping.ColumnData.TargetEntityProperty)
            .Where(group => group.Count() > 1)
            .Select(group => group.Key)
            .ToList();

        if(duplicateProperties.Any())
        {
            throw new InvalidOperationException($"Duplicate TargetEntityProperty detected in column mappings: {string.Join(", ", duplicateProperties)}");
        }
    }

    #endregion
    
    #endregion
    
    #region Date Assertions

    private bool IsValidDateRange((DateTime startDate, DateTime? endDate) dateRange)
    {
        return !dateRange.endDate.HasValue || dateRange.endDate >= dateRange.startDate;
    }

    private bool IsNewDateRangeContainedInExistingDateRange(
        (DateTime startDate, DateTime? endDate) newDateRange, 
        (DateTime startDate, DateTime? endDate) existingDateRange)
    {
        if(newDateRange.endDate.HasValue)
        {
            return newDateRange.startDate >= existingDateRange.startDate &&
                   (existingDateRange.endDate == null || newDateRange.endDate <= existingDateRange.endDate);
        }

        return newDateRange.startDate >= existingDateRange.startDate &&
               (existingDateRange.endDate == null || newDateRange.startDate <= existingDateRange.endDate);
    }

    private bool IsNewDateRangeOverlappingExistingDateRanges(
        (DateTime startDate, DateTime? endDate) newDateRange, 
        List<(DateTime startDate, DateTime? endDate)> existingDateRanges, 
        bool ignoreExactMatch = false)
    {
        foreach((DateTime startDate, DateTime? endDate) existing in existingDateRanges)
        {
            DateTime existingEnd = existing.endDate ?? DateTime.MaxValue;
            DateTime newEnd = newDateRange.endDate ?? DateTime.MaxValue;

            bool isExactMatch =
                existing.startDate == newDateRange.startDate &&
                existingEnd == newEnd;

            if(ignoreExactMatch && isExactMatch)
                continue;

            bool overlaps =
                newDateRange.startDate < existingEnd &&
                existing.startDate < newEnd;

            if(overlaps)
                return true;
        }

        return false;
    }

    #endregion
    
    #region Misc Static
    
    private static readonly HashSet<string> TabularFileImportModelPropertyNames =
        typeof(TabularFileImportModel)
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Select(p => p.Name)
            .ToHashSet();
    
    #endregion
    
    #endregion
    
    #region Export
    
    #region Processing
    
    private async Task ProcessSheets(
        WorkbookPart workbookPart,
        Sheets sheets,
        string sheetNameRaw,
        TabularFileSheetMapping mapping,
        uint sheetId,
        Func<Type, CancellationToken, Task<IEnumerable<object>>> fetchEntitiesAsync,
        TabularFileExportResult result,
        CancellationToken cancellationToken)
    {
        string sheetName = CleanSheetName(sheetNameRaw);

        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        SheetData sheetData = new();
        worksheetPart.Worksheet = new Worksheet(sheetData);

        Sheet sheet = new()
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = sheetName
        };
        sheets.Append(sheet);

        IReadOnlyList<KeyValuePair<string, TabularFileColumnMapping>> columnMappings = mapping.ColumnMappings.ToList();

        AppendHeaderRow(sheetData, columnMappings);

        IEnumerable<object> entities = await fetchEntitiesAsync(mapping.EntityType, cancellationToken);

        AppendDataRows(sheetData, columnMappings, entities, mapping.EntityType, result);
    }
    
    private static void AppendHeaderRow(
        SheetData sheetData, 
        IReadOnlyList<KeyValuePair<string, TabularFileColumnMapping>> columnMappings)
    {
        Row headerRow = new() { RowIndex = 1 };
        uint currentColumnIndex = 1;

        foreach (KeyValuePair<string, TabularFileColumnMapping> colKvp in columnMappings)
        {
            TabularFileColumnMapping colMap = colKvp.Value;
            string headerText = !string.IsNullOrWhiteSpace(colMap.OriginalHeaderName)
                ? colMap.OriginalHeaderName
                : colKvp.Key;

            Cell headerCell = CreateTextCell(GetColumnName(currentColumnIndex), 1U, headerText);
            headerCell.StyleIndex = 2U;
            headerRow.Append(headerCell);
            currentColumnIndex++;
        }
        sheetData.Append(headerRow);
    }
    
    private void AppendDataRows(
        SheetData sheetData,
        IReadOnlyList<KeyValuePair<string, TabularFileColumnMapping>> columnMappings,
        IEnumerable<object> entities,
        Type entityType,
        TabularFileExportResult result)
    {
        uint dataRowIndex = 2;

        foreach (object entity in entities)
        {
            object? importModel = TryGetImportModel(entity, out string? importModelError);

            if (importModel is null)
            {
                if (importModelError != null)
                    result.Warnings.Add($"Entity of type {entityType.Name} at export: {importModelError}. Writing empty row.");

                Row emptyRow = new() { RowIndex = dataRowIndex };
                for (uint i = 1; i <= columnMappings.Count; i++)
                    emptyRow.Append(CreateTextCell(GetColumnName(i), dataRowIndex, string.Empty));

                sheetData.Append(emptyRow);
                dataRowIndex++;
                continue;
            }

            AppendDataRow(sheetData, columnMappings, importModel, dataRowIndex, result);
            dataRowIndex++;
        }
    }
    
    private void AppendDataRow(
        SheetData sheetData,
        IReadOnlyList<KeyValuePair<string, TabularFileColumnMapping>> columnMappings,
        object importModel,
        uint dataRowIndex,
        TabularFileExportResult result)
    {
        Row dataRow = new() { RowIndex = dataRowIndex };
        uint currentColumnIndex = 1;

        foreach (KeyValuePair<string, TabularFileColumnMapping> colKvp in columnMappings)
        {
            TabularFileColumnMapping colMap = colKvp.Value;
            string targetProperty = colMap.ColumnData.TargetEntityProperty;

            object? propValue = null;
            PropertyInfo? propInfo = importModel.GetType().GetProperty(targetProperty, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
            if (propInfo is not null)
            {
                propValue = propInfo.GetValue(importModel);
            }
            else
            {
                result.Warnings.Add($"Property '{targetProperty}' not found on import model '{importModel.GetType().Name}'.");
            }

            Cell cell;

            if (colMap.Stringifier is null)
            {
                cell = CreateNativeValueCell(propValue, GetColumnName(currentColumnIndex), dataRowIndex);
            }
            else
            {
                StringifierResult stringifierResult = colMap.Stringifier(propValue);

                cell = CreateTextCell(
                    header: GetColumnName(currentColumnIndex),
                    rowIndex: dataRowIndex,
                    text: stringifierResult.IsValid
                        ? stringifierResult.StringValue ?? string.Empty
                        : string.Empty);

                if (!stringifierResult.IsValid)
                    result.Warnings.Add($"Stringifier failed for column '{colKvp.Key}' (property '{targetProperty}') on row {dataRowIndex}. Writing empty cell.");
            }

            dataRow.Append(cell);
            currentColumnIndex++;
        }

        sheetData.Append(dataRow);
    }
    
    private static object? TryGetImportModel(object? entity, out string? error)
    {
        error = null;
        switch (entity)
        {
            case null: error = "entity is null"; return null;
            case TabularFileImportModel: return entity;
        }

        const string methodName = nameof(IAppItemToImportModel<>.ToImportModel);
        
        MethodInfo? toImportMethod = entity
            .GetType()
            .GetMethod(methodName, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
        
        if (toImportMethod is null)
        {
            error = $"No {methodName} method found on type {entity.GetType().Name}";
            return null;
        }

        try
        {
            return toImportMethod.Invoke(entity, []);
        }
        catch (Exception ex)
        {
            error = $"{methodName} invocation failed: {ex.Message}";
            return null;
        }
    }
    
    #endregion
    
    #region Formating
    
    private static string CleanSheetName(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return "Sheet";
        string cleaned = Regex.Replace(raw, @"[\[\]\:\*\?\/\\]", string.Empty);
        
        if (cleaned.Length > TabellaConstants.ExcelSheetNameMaxCharacterCount)
            cleaned = cleaned[..TabellaConstants.ExcelSheetNameMaxCharacterCount];
        
        return cleaned;
    }
    
    private static string GetColumnName(uint index)
    {
        uint dividend = index;
        string columnName = string.Empty;
        while (dividend > 0)
        {
            int modulo = (int)((dividend - 1) % 26);
            columnName = Convert.ToChar('A' + modulo) + columnName;
            dividend = (dividend - 1) / 26;
        }
        return columnName;
    }
    
    private Stylesheet CreateStylesheet()
    {
        const uint dateFormatId = (uint)ExcelStyleFormatsEnum.DateTime;
        
        Fonts fonts = new(
            new Font(),                      
            new Font(new Bold()),
            new Font(                       
                new Bold(),
                new Color { Rgb = HexBinaryValue.FromString(options.Value.ExcelHeaderRowTextColor) } 
            )
        );

        Fills fills = new(
            new Fill(new PatternFill { PatternType = PatternValues.None }),    
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 }),   
            new Fill(                                                            
                new PatternFill(
                    new ForegroundColor { Rgb = HexBinaryValue.FromString(options.Value.ExcelHeaderRowBackgroundFillColor) }
                ) { PatternType = PatternValues.Solid }
            )
        );

        Borders borders = new(new Border()); 

        CellFormats cellFormats = new(
            new CellFormat(), 
            new CellFormat   
            {
                NumberFormatId = dateFormatId,
                ApplyNumberFormat = true
            },
            new CellFormat   
            {
                FontId = 2,
                FillId = 2,
                ApplyFont = true,
                ApplyFill = true,
                ApplyNumberFormat = false  
            }
        );

        return new Stylesheet
        {
            Fonts = fonts,
            Fills = fills,
            Borders = borders,
            CellFormats = cellFormats
        };
    }
    
    #endregion
    
    #region Cell Creation
    
    private static Cell CreateTextCell(string header, uint rowIndex, string text)
    {
        Cell cell = new()
        {
            CellReference = $"{header}{rowIndex}",
            CellValue = new CellValue { Text = text },
            DataType = CellValues.InlineString
        };
        
        InlineString inlineString = new();
        Text t = new() { Text = text };
        inlineString.AppendChild(t);
        cell.AppendChild(inlineString);

        return cell;
    }
    
    private Cell CreateNativeValueCell(object? value, string columnName, uint rowIndex)
    {
        if (value is null)
            return CreateTextCell(columnName, rowIndex, string.Empty);

        return value switch
        {
            int i => CreateNumberCell(columnName, rowIndex, i.ToString()),
            long l => CreateNumberCell(columnName, rowIndex, l.ToString()),
            float f => CreateNumberCell(columnName, rowIndex, f.ToString(CultureInfo.InvariantCulture)),
            double d => CreateNumberCell(columnName, rowIndex, d.ToString(CultureInfo.InvariantCulture)),
            decimal m => CreateNumberCell(columnName, rowIndex, m.ToString(CultureInfo.InvariantCulture)),
        
            bool b => CreateBooleanCell(columnName, rowIndex, b),
        
            DateTime dt => CreateExcelDateCell(columnName, rowIndex, dt),
        
            _ => CreateTextCell(columnName, rowIndex, value.ToString() ?? string.Empty)
        };
    }
    
    private Cell CreateNumberCell(string col, uint row, string value)
    {
        return new Cell
        {
            CellReference = col + row,
            CellValue = new CellValue(value),
            DataType = CellValues.Number
        };
    }
    
    private Cell CreateBooleanCell(string col, uint row, bool value)
    {
        return new Cell
        {
            CellReference = col + row,
            DataType = CellValues.Boolean,
            CellValue = new CellValue(value ? "1" : "0")
        };
    }
    
    private Cell CreateExcelDateCell(string col, uint row, DateTime date)
    {
        double oa = date.ToOADate();

        return new Cell
        {
            CellReference = col + row,
            CellValue = new CellValue(oa.ToString(CultureInfo.InvariantCulture)),
            DataType = CellValues.Number,
            StyleIndex = 1u
        };
    }
    
    #endregion
        
    #endregion
}