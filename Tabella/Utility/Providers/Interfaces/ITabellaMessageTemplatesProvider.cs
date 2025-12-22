using MessageKit.Data.Core;

namespace Tabella.Utility.Providers.Interfaces;

/// <summary>
/// Interface for providing Tabella-specific message templates.
/// </summary>
public interface ITabellaMessageTemplatesProvider
{
    #region Import

    /// <summary>
    /// Message template for when no worksheets are found in the imported file.
    /// </summary>
    MessageTemplate NoWorksheetsFound { get; }
    /// <summary>
    /// Message template for when a specified worksheet is not found in the imported file.
    /// </summary>
    MessageTemplate WorksheetNotFound { get; }
    /// <summary>
    /// Message template for when a header is not found at the expected index in the imported file.
    /// </summary>
    MessageTemplate HeaderNotFoundAtIndex { get; }
    /// <summary>
    /// Message template for when a required entity cannot be deleted because it is in use.
    /// </summary>
    MessageTemplate CannotDeleteEntity { get; }
    /// <summary>
    /// Message template for when a duplicate entry is found during import.
    /// </summary>
    MessageTemplate DuplicateEntry { get; }
    /// <summary>
    /// Message template for when an invalid date range is encountered.
    /// </summary>
    MessageTemplate InvalidDateRange { get; }
    /// <summary>
    /// Message template for when a date range overlaps with an existing date range.
    /// </summary>
    MessageTemplate DateRangeOverlapsExisting { get; }
    /// <summary>
    /// Message template for when an active entity cannot be modified.
    /// </summary>
    MessageTemplate CannotModifyActiveEntity { get; }
    /// <summary>
    /// Message template for when there is an error processing a cell during import.
    /// </summary>
    MessageTemplate CellProcessingError { get; }
    /// <summary>
    /// Message template for when a required field is missing during import.
    /// </summary>
    MessageTemplate RequiredField { get; }
    /// <summary>
    /// Message template for when no accompanying item is found during import.
    /// </summary>
    MessageTemplate NoAccompanyingItemFound { get; }
    /// <summary>
    /// Message template for when a specified column is not found in the imported file.
    /// </summary>
    MessageTemplate ColumnNotFound { get; }
    /// <summary>
    /// Message template for when an incorrect column is found in the imported file.
    /// </summary>
    MessageTemplate WrongColumnInFile { get; }
    /// <summary>
    /// Message template for import warning when no data was imported.
    /// </summary>
    MessageTemplate ImportWarningNoDataImported { get; }
    /// <summary>
    /// Message template for import warning when some data was not imported.
    /// </summary>
    MessageTemplate ImportWarningSomeDataNotImported { get; }

    #endregion

    #region Custom Cast

    /// <summary>
    /// Message template for when a generic cast fails.
    /// </summary>
    MessageTemplate CastMessageGeneric { get; }
    /// <summary>
    /// Message template for when a date cast fails.
    /// </summary>
    MessageTemplate CastMessageDate { get; }
    /// <summary>
    /// Message template for when a boolean cast fails.
    /// </summary>
    MessageTemplate CastMessageBool { get; }

    #endregion

    #region Custom Validation

    /// <summary>
    /// Message template for when a null value is not allowed.
    /// </summary>
    MessageTemplate ValidationMessageNullNotAllowed { get; }
    /// <summary>
    /// Message template for when a numeric value is out of the allowed range.
    /// </summary>
    MessageTemplate ValidationMessageInvalidNumericRange { get; }
    /// <summary>
    /// Message template for when a string's length is invalid.
    /// </summary>
    MessageTemplate ValidationMessageInvalidStringType { get; }
    /// <summary>
    /// Message template for when a string's length is outside the allowed range.
    /// </summary>
    MessageTemplate ValidationMessageInvalidStringLength { get; }
    /// <summary>
    /// Message template for when a decimal value is of an invalid type.
    /// </summary>
    MessageTemplate ValidationMessageInvalidDecimalType { get; }
    /// <summary>
    /// Message template for when a decimal value has invalid precision or scale.
    /// </summary>
    MessageTemplate ValidationMessageInvalidDecimalPrecisionScale { get; }

    #endregion
}