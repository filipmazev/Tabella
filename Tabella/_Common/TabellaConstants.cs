using Tabella.Enums;

namespace Tabella._Common;

/// <summary>
/// Contains constant values used throughout the Tabella application.
/// </summary>
public static class TabellaConstants
{
    #region Limits

    /// <summary>
    /// Maximum number of characters allowed in an Excel sheet name.
    /// </summary>
    public const int ExcelSheetNameMaxCharacterCount = 31;

    #endregion
    
    #region Formats
    
    /// <summary>
    /// Standard decimal format with up to 10 decimal places.
    /// </summary>
    public const string DecimalFormat = "0.##########";
    /// <summary>
    /// Standard date format (dd-MM-yyyy).
    /// </summary>
    public const string DateFormat = "dd-MM-yyyy";
    /// <summary>
    /// Standard date and time format (dd-MM-yyyy HH:mm:ss).
    /// </summary>
    public const string DateTimeFormat = "dd-MM-yyyy HH:mm:ss";
    /// <summary>
    /// Standard time format (HH:mm:ss).
    /// </summary>
    public const string StringBoolTrue = "true";
    /// <summary>
    /// Standard string representation for boolean false.
    /// </summary>
    public const string StringBoolFalse = "false";
    
    #endregion
    
    #region Misc
    
    /// <summary>
    /// Representation of an empty field in imports/exports.
    /// </summary>
    public const string EmptyField = "/";
    /// <summary>
    /// Representation of a not assigned value in imports/exports.
    /// </summary>
    public const string NotAssigned = "N/A";
    /// <summary>
    /// Comma separator used in CSV files.
    /// </summary>
    public const string CommaSeparator = ",";
    /// <summary>
    /// Semicolon separator used in CSV files.
    /// </summary>
    public const string SemicolonSeparator = ";";
    /// <summary>
    /// Tab separator used in TSV files.
    /// </summary>
    public const string SheetedFileImportedEntityValidatorSingleGroupingKey = "_GROUPING_KEY_";
    
    #endregion

    #region Message Placeholders
    
    /// <summary>
    /// Symbol representing no end in a range.
    /// </summary>
    public const string MessageRangeNoEndSymbol = "\u221e";

    /// <summary>
    /// Placeholder for the value that was attempted to be cast.
    /// </summary>
    public const string MessagePlaceholderCastedValue = "CastedValue";
    /// <summary>
    /// Placeholder for the current value being processed.
    /// </summary>
    public const string MessagePlaceholderCurrentValue = "CurrentValue";
    /// <summary>
    /// Placeholder for item information in messages.
    /// </summary>
    public const string MessagePlaceholderItemInformation = "ItemInformation";
    /// <summary>
    /// Placeholder for the index in messages.
    /// </summary>
    public const string MessagePlaceholderIndex = "Index";

    /// <summary>
    /// Placeholder for the start date in messages.
    /// </summary>
    public const string MessagePlaceholderStartDate = "StartDate";
    
    /// <summary>
    /// Placeholder for the end date in messages.
    /// </summary>
    public const string MessagePlaceholderEndDate = "EndDate";

    #endregion
    
    #region Excel
    
    /// <summary>
    /// The base date for Excel date calculations (January 1, 1900).
    /// </summary>
    public static readonly DateTime ExcelBaseDate = new DateTime(1900, 1, 1);
    
    /// <summary>
    /// The number of days to adjust for the Excel leap year bug.
    /// </summary>
    public const int ExcelLeapYearBug = 59;

    /// <summary>
    /// List of standard OpenXML date format IDs.
    /// </summary>
    public static readonly List<uint> OpenXmlStandardDateFormatIds =
    [
        (uint)ExcelStyleFormatsEnum.DateShort,
        (uint)ExcelStyleFormatsEnum.DateTime
    ];

    /// <summary>
    /// List of recognized OpenXML date formats.
    /// </summary>
    public static readonly List<string> OpenXmlDateFormats =
    [
        "m", "M", "mm", "MM", "mmm", "MMM", "mmmm", "MMMM", "mmmmm", "MMMMM",
        "d", "D", "dd", "DD", "ddd", "DDD", "dddd", "DDDD",
        "yy", "YY", "yyyy", "YYYY",

        "MM/dd/yyyy", "MM/DD/YYYY", "mm/dd/yyyy", "mm/DD/yyyy", "MM/dd/yy", "MM/DD/YY", "mm/dd/yy", "mm/DD/yy",
        "MM-dd-yyyy", "MM-DD-YYYY", "mm-dd-yyyy", "mm-DD-yyyy", "MM-dd-yy", "MM-DD-YY", "mm-dd-yy", "mm-DD-yy",
        "MM.dd.yyyy", "MM.DD.YYYY", "mm.dd.yyyy", "mm.DD.yyyy", "MM.dd.yy", "MM.DD.YY", "mm.dd.yy", "mm.DD.yy",

        "dd/MM/yyyy", "DD/MM/YYYY", "dd/mm/yyyy", "DD/mm/yyyy", "dd/MM/yy", "DD/MM/YY", "dd/mm/yy", "DD/mm/yy",
        "dd-MM-yyyy", "DD-MM-YYYY", "dd-mm-yyyy", "DD-mm-yyyy", "dd-MM-yy", "DD-MM-YY", "dd-mm-yy", "DD-mm-yy",
        "dd.MM.yyyy", "DD.MM.YYYY", "dd.mm.yyyy", "DD.mm.yyyy", "dd.MM.yy", "DD.MM.YY", "dd.mm.yy", "DD.mm.yy",

        "yyyy/MM/dd", "YYYY/MM/DD", "yyyy/mm/dd", "YYYY/mm/dd",
        "yyyy-MM-dd", "YYYY-MM-DD", "yyyy-mm-dd", "YYYY-mm-dd",
        "yyyy.MM.dd", "YYYY.MM.DD", "yyyy.mm.dd", "YYYY.mm.dd",

        "yy/MM/dd", "YY/MM/DD", "yy/mm/dd", "YY/mm/dd",
        "yy-MM-dd", "YY-MM-DD", "yy-mm-dd", "YY-mm-dd",
        "yy.MM.dd", "YY.MM.DD", "yy.mm.dd", "YY.mm.dd",

        "d/m/yyyy", "D/M/YYYY", "d/M/yyyy", "D/m/YYYY",
        "d-m-yyyy", "D-M-YYYY", "d-M-yyyy", "D-m-YYYY",
        "d.m.yyyy", "D.M.YYYY", "d.M.yyyy", "D.m.YYYY",

        "m/d/yyyy", "M/D/YYYY", "m/D/yyyy", "M/d/YYYY",
        "m-d-yyyy", "M-D-YYYY", "m-D-yyyy", "M-d-YYYY",
        "m.d.yyyy", "M.D.YYYY", "m.D.yyyy", "M.d.YYYY",

        "d/m/yy", "D/M/YY", "d/M/yy", "D/m/YY",
        "d-m-yy", "D-M-YY", "d-M-yy", "D-m-YY",
        "d.m.yy", "D.M.YY", "d.M.yy", "D.m.YY",

        "m/d/yy", "M/D/YY", "m/D/yy", "M/d/YY",
        "m-d-yy", "M-D-YY", "m-D-yy", "M-d-YY",
        "m.d.yy", "M.D.YY", "m.D.yy", "M.d.YY",

        "MM/dd/yyyy h:mm AM/PM", "MM/DD/YYYY H:MM AM/PM",
        "MM/dd/yyyy hh:mm:ss", "MM/DD/YYYY HH:MM:SS",
        "yyyy-MM-dd HH:mm:ss", "YYYY-MM-DD HH:MM:SS",
        "dd-MM-yyyy HH:mm:ss", "DD-MM-YYYY HH:MM:SS",
        "d/m/yyyy H:mm:ss", "D/M/YYYY H:MM:SS",
        "m/d/yyyy h:mm:ss A/P", "M/D/YYYY H:MM:SS A/P",

        "r", "rr", "g", "gg", "ggg", "e", "ee",
        "b1", "b2",

        "0xf800", "0xf400"
    ];
    
    #endregion
}