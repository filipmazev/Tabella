namespace Tabella.Configuration;

/// <summary>
/// Configuration options for tabular file import messages.
/// </summary>
public class TabularFileImportMessageOptions
{
    /// <summary>
    /// The default sender name for messages.
    /// </summary>
    public string DefaultSender { get; set; } = "System";
    
    /// <summary>
    /// The background fill color for Excel header rows in ARGB hex format.
    /// </summary>
    public string ExcelHeaderRowBackgroundFillColor { get; set; } = "FFF54900";
    
    /// <summary>
    /// The text color for Excel header rows in ARGB hex format.
    /// </summary>
    public string ExcelHeaderRowTextColor  { get; set; } = "FFFFFFFF";
    
    #region Templates
    
    #region Import
    
    /// <summary>
    /// Key for the message indicating no worksheets were found in the imported file.
    /// </summary>
    public string NoWorksheetsFoundKey { get; set; } = "messages.no_worksheets_found";

    /// <summary>
    /// Key for the message indicating a specific worksheet was not found.
    /// </summary>
    public string WorksheetNotFoundKey { get; set; } = "messages.worksheet_not_found";

    /// <summary>
    /// Key for the message indicating the header row was not found at the expected index.
    /// </summary>
    public string HeaderNotFoundAtIndexKey { get; set; } = "messages.sheet_processing_message_header_row_not_found_at_index";

    /// <summary>
    /// Key for the message indicating a required entity cannot be deleted because it is in use.
    /// </summary>
    public string CannotDeleteEntityKey { get; set; } = "messages.sheet_processing_message_cannot_delete_entity";

    /// <summary>
    /// Key for the message indicating a duplicate entry was found during import.
    /// </summary>
    public string DuplicateEntryKey { get; set; } = "messages.sheet_processing_message_duplicate";

    /// <summary>
    /// Key for the message indicating an invalid date range was encountered.
    /// </summary>
    public string InvalidDateRangeKey { get; set; } = "messages.sheet_processing_message_invalid_date_range";

    /// <summary>
    /// Key for the message indicating a date range overlaps with an existing date range.
    /// </summary>
    public string DateRangeOverlapsExistingKey { get; set; } = "messages.sheet_processing_message_date_range_overlaps_existing";

    /// <summary>
    /// Key for the message indicating an active entity cannot be modified.
    /// </summary>
    public string CannotModifyActiveEntityKey { get; set; } = "messages.sheet_processing_message_cannot_modify_active_entity";

    /// <summary>
    /// Key for the message indicating there was an error processing a cell during import.
    /// </summary>
    public string CellProcessingErrorKey { get; set; } = "messages.sheet_cell_processing_error";

    /// <summary>
    /// Key for the message indicating a required field is missing during import.
    /// </summary>
    public string RequiredFieldKey { get; set; } = "messages.sheet_processing_message_required_field";

    /// <summary>
    /// Key for the message indicating no accompanying item was found during import.
    /// </summary>
    public string NoAccompanyingItemFoundKey { get; set; } = "messages.sheet_processing_message_no_accompanying_item_found";

    /// <summary>
    /// Key for the message indicating a specified column was not found in the imported file.
    /// </summary>
    public string ColumnNotFoundKey { get; set; } = "messages.sheet_processing_message_column_not_found";

    /// <summary>
    /// Key for the message indicating an incorrect column was found in the imported file.
    /// </summary>
    public string WrongColumnInFileKey { get; set; } = "messages.sheet_processing_message_column_is_wrong";

    /// <summary>
    /// Key for the import warning message when no data was imported.
    /// </summary>
    public string ImportWarningNoDataImportedKey { get; set; } = "messages.import_info_message_no_data_imported";

    /// <summary>
    /// Key for the import warning message when some data was not imported.
    /// </summary>
    public string ImportWarningSomeDataNotImportedKey { get; set; } = "messages.import_info_message_some_data_not_imported";

    #endregion
    
    #region Custom Cast
    
    /// <summary>
    /// Key for the message indicating a generic cast failure.
    /// </summary>
    public string CastMessageGenericKey { get; set; } = "messages.cast_message_generic";

    /// <summary>
    /// Key for the message indicating a date cast failure.
    /// </summary>
    public string CastMessageDateKey { get; set; } = "messages.cast_message_date";

    /// <summary>
    /// Key for the message indicating a boolean cast failure.
    /// </summary>
    public string CastMessageBoolKey { get; set; } = "messages.cast_message_bool";

    #endregion
    
    #region Custom Validation
    
    /// <summary>
    /// Key for the message indicating that a null value is not allowed.
    /// </summary>
    public string ValidationMessageNullNotAllowedKey { get; set; } = "messages.value_validation_message_null_not_allowed";

    /// <summary>
    /// Key for the message indicating that a numeric value is outside the allowed range.
    /// </summary>
    public string ValidationMessageInvalidNumericRangeKey { get; set; } = "messages.value_validation_message_invalid_numeric_range";

    /// <summary>
    /// Key for the message indicating that a string value is of an invalid type.
    /// </summary>
    public string ValidationMessageInvalidStringTypeKey { get; set; } = "messages.value_validation_message_invalid_string_type";

    /// <summary>
    /// Key for the message indicating that a string's length is outside the allowed range.
    /// </summary>
    public string ValidationMessageInvalidStringLengthKey { get; set; } = "messages.value_validation_message_invalid_string_length";

    /// <summary>
    /// Key for the message indicating that a decimal value is of an invalid type.
    /// </summary>
    public string ValidationMessageInvalidDecimalTypeKey { get; set; } = "messages.value_validation_message_invalid_decimal_type";

    /// <summary>
    /// Key for the message indicating that a decimal value has invalid precision or scale.
    /// </summary>
    public string ValidationMessageInvalidDecimalPrecisionScaleKey { get; set; } = "messages.value_validation_message_invalid_decimal_precision_scale";
    
    #endregion
    
    #endregion
}