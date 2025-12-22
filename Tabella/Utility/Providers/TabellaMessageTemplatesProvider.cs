using Microsoft.Extensions.Options;
using Tabella.Configuration;
using MessageKit.Data.Core;
using Tabella.Data.Core;
using Tabella._Common;
using Tabella.Utility.Providers.Interfaces;

namespace Tabella.Utility.Providers;

/// <summary>
/// Provider for tabular file import message templates.
/// </summary>
public class TabellaMessageTemplatesProvider : ITabellaMessageTemplatesProvider
{
    #region Import

    /// <summary>
    /// Message template for when no worksheets are found in the file.
    /// </summary>
    public MessageTemplate NoWorksheetsFound { get; }
    /// <summary>
    /// Message template for when a specified worksheet is not found in the file.
    /// </summary>
    public MessageTemplate WorksheetNotFound { get; }
    /// <summary>
    /// Message template for when a header is not found at the expected index.
    /// </summary>
    public MessageTemplate HeaderNotFoundAtIndex { get; }
    /// <summary>
    /// Message template for when a required entity cannot be deleted because it is in use.
    /// </summary>
    public MessageTemplate CannotDeleteEntity { get; }
    /// <summary>
    /// Message template for when a duplicate entry is found during import.
    /// </summary>
    public MessageTemplate DuplicateEntry { get; }
    /// <summary>
    /// Message template for when an invalid date range is encountered.
    /// </summary>
    public MessageTemplate InvalidDateRange { get; }
    /// <summary>
    /// Message template for when a date range overlaps with an existing date range.
    /// </summary>
    public MessageTemplate DateRangeOverlapsExisting { get; }
    /// <summary>
    /// Message template for when an active entity cannot be modified.
    /// </summary>
    public MessageTemplate CannotModifyActiveEntity { get; }
    /// <summary>
    /// Message template for when there is an error processing a cell during import.
    /// </summary>
    public MessageTemplate CellProcessingError { get; }
    /// <summary>
    /// Message template for when a required field is missing during import.
    /// </summary>
    public MessageTemplate RequiredField { get; }
    /// <summary>
    /// Message template for when no accompanying item is found during import.
    /// </summary>
    public MessageTemplate NoAccompanyingItemFound { get; }
    /// <summary>
    /// Message template for when a specified column is not found in the file.
    /// </summary>
    public MessageTemplate ColumnNotFound { get; }
    /// <summary>
    /// Message template for when an incorrect column is found in the file.
    /// </summary>
    public MessageTemplate WrongColumnInFile { get; }
    /// <summary>
    /// Message template for when no data is imported.
    /// </summary>
    public MessageTemplate ImportWarningNoDataImported { get; }
    /// <summary>
    /// Message template for when some data is not imported.
    /// </summary>
    public MessageTemplate ImportWarningSomeDataNotImported { get; }

    #endregion

    #region Custom Cast

    /// <summary>
    /// Message template for generic casting errors.
    /// </summary>
    public MessageTemplate CastMessageGeneric { get; }
    /// <summary>
    /// Message template for date casting errors.
    /// </summary>
    public MessageTemplate CastMessageDate { get; }
    /// <summary>
    /// Message template for boolean casting errors.
    /// </summary>
    public MessageTemplate CastMessageBool { get; }

    #endregion

    #region Custom Validation

    /// <summary>
    /// Message template for when a null value is not allowed.
    /// </summary>
    public MessageTemplate ValidationMessageNullNotAllowed { get; }
    /// <summary>
    /// Message template for when a numeric value is outside the allowed range.
    /// </summary>
    public MessageTemplate ValidationMessageInvalidNumericRange { get; }
    /// <summary>
    /// Message template for when a string value is of an invalid type.
    /// </summary>
    public MessageTemplate ValidationMessageInvalidStringType { get; }
    /// <summary>
    /// Message template for when a string's length is outside the allowed range.
    /// </summary>
    public MessageTemplate ValidationMessageInvalidStringLength { get; }
    /// <summary>
    /// Message template for when a decimal value is of an invalid type.
    /// </summary>
    public MessageTemplate ValidationMessageInvalidDecimalType { get; }
    /// <summary>
    /// Message template for when a decimal value has invalid precision or scale.
    /// </summary>
    public MessageTemplate ValidationMessageInvalidDecimalPrecisionScale { get; }

    #endregion

    /// <summary>
    /// Constructor for TabellaMessageTemplatesProvider
    /// </summary>
    /// <param name="options"></param>
    public TabellaMessageTemplatesProvider(IOptions<TabularFileImportMessageOptions> options)
    {
        TabularFileImportMessageOptions opts = options.Value;

        #region Import

        NoWorksheetsFound = new MessageTemplate(opts.NoWorksheetsFoundKey);

        WorksheetNotFound = new MessageTemplate(
            opts.WorksheetNotFoundKey,
            requiredPlaceholders:
            [
                nameof(TabularFileProcessedObject.SheetName)
            ]);

        HeaderNotFoundAtIndex = new MessageTemplate(
            opts.HeaderNotFoundAtIndexKey,
            requiredPlaceholders:
            [
                TabellaConstants.MessagePlaceholderIndex
            ]);

        CannotDeleteEntity = new MessageTemplate(
            opts.CannotDeleteEntityKey,
            requiredPlaceholders:
            [
                TabellaConstants.MessagePlaceholderIndex,
                TabellaConstants.MessagePlaceholderItemInformation
            ]);

        DuplicateEntry = new MessageTemplate(opts.DuplicateEntryKey);

        InvalidDateRange = new MessageTemplate(
            opts.InvalidDateRangeKey,
            requiredPlaceholders:
            [
                TabellaConstants.MessagePlaceholderStartDate,
                TabellaConstants.MessagePlaceholderEndDate
            ]);

        DateRangeOverlapsExisting = new MessageTemplate(
            opts.DateRangeOverlapsExistingKey,
            requiredPlaceholders:
            [
                TabellaConstants.MessagePlaceholderStartDate,
                TabellaConstants.MessagePlaceholderEndDate
            ]);

        CannotModifyActiveEntity = new MessageTemplate(
            opts.CannotModifyActiveEntityKey,
            requiredPlaceholders:
            [
                TabellaConstants.MessagePlaceholderItemInformation
            ]);

        CellProcessingError = new MessageTemplate(opts.CellProcessingErrorKey);

        RequiredField = new MessageTemplate(
            opts.RequiredFieldKey,
            requiredPlaceholders:
            [
                nameof(TabularFileCellData.ColumnName)
            ]);

        NoAccompanyingItemFound = new MessageTemplate(
            opts.NoAccompanyingItemFoundKey,
            requiredPlaceholders:
            [
                TabellaConstants.MessagePlaceholderCurrentValue
            ]);

        ColumnNotFound = new MessageTemplate(
            opts.ColumnNotFoundKey,
            requiredPlaceholders:
            [
                nameof(TabularFileCellData.ColumnName)
            ]);

        WrongColumnInFile = new MessageTemplate(
            opts.WrongColumnInFileKey,
            requiredPlaceholders:
            [
                nameof(TabularFileCellData.ColumnName)
            ]);

        ImportWarningNoDataImported = new MessageTemplate(opts.ImportWarningNoDataImportedKey);

        ImportWarningSomeDataNotImported = new MessageTemplate(opts.ImportWarningSomeDataNotImportedKey);

        #endregion

        #region Custom Cast

        CastMessageGeneric = new MessageTemplate(
            opts.CastMessageGenericKey,
            requiredPlaceholders:
            [
                TabellaConstants.MessagePlaceholderCastedValue
            ]);

        CastMessageDate = new MessageTemplate(
            opts.CastMessageDateKey,
            requiredPlaceholders:
            [
                TabellaConstants.MessagePlaceholderCastedValue
            ]);

        CastMessageBool = new MessageTemplate(
            opts.CastMessageBoolKey,
            requiredPlaceholders:
            [
                TabellaConstants.MessagePlaceholderCastedValue
            ]);

        #endregion

        #region Custom Validation

        ValidationMessageNullNotAllowed = new MessageTemplate(opts.ValidationMessageNullNotAllowedKey);

        ValidationMessageInvalidNumericRange = new MessageTemplate(
            opts.ValidationMessageInvalidNumericRangeKey,
            requiredPlaceholders:
            [
                TabellaConstants.MessagePlaceholderCurrentValue,
                "Min",
                "Max"
            ]);

        ValidationMessageInvalidStringType = new MessageTemplate(
            opts.ValidationMessageInvalidStringTypeKey,
            requiredPlaceholders:
            [
                TabellaConstants.MessagePlaceholderCurrentValue
            ]);

        ValidationMessageInvalidStringLength = new MessageTemplate(
            opts.ValidationMessageInvalidStringLengthKey,
            requiredPlaceholders:
            [
                TabellaConstants.MessagePlaceholderCurrentValue,
                "MinLength",
                "MaxLength"
            ]);

        ValidationMessageInvalidDecimalType = new MessageTemplate(
            opts.ValidationMessageInvalidDecimalTypeKey,
            requiredPlaceholders:
            [
                TabellaConstants.MessagePlaceholderCurrentValue
            ]);

        ValidationMessageInvalidDecimalPrecisionScale = new MessageTemplate(
            opts.ValidationMessageInvalidDecimalPrecisionScaleKey,
            requiredPlaceholders:
            [
                TabellaConstants.MessagePlaceholderCurrentValue,
                "Precision",
                "Scale"
            ]);

        #endregion
    }
}