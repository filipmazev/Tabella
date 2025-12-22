using Tabella.Utility.Delegates;

namespace Tabella.Data.Core;

/// <summary>
/// Represents the mapping configuration for a column in a tabular file.
/// </summary>
public class TabularFileColumnMapping
{
    /// <summary>
    /// The set of acceptable cell types for this column.
    /// </summary>
    public required HashSet<Type> CellTypes { get; set; }
    
    /// <summary>
    /// The data associated with this column.
    /// </summary>
    public required TabularFileColumnData ColumnData { get; set; }
    
    /// <summary>
    /// Custom casting logic for converting cell values to the desired type.
    /// </summary>
    public CustomCastDelegate? CustomCast { get; set; }
    
    /// <summary>
    /// Custom validation logic for validating cell values.
    /// </summary>
    public List<CustomValidatorDelegate> CustomValidators { get; } = [];
    
    /// <summary>
    /// Custom stringification logic for converting cell values to strings.
    /// </summary>
    public StringifierDelegate? Stringifier { get; set; }
    
    /// <summary>
    /// Indicates whether the column was found in the file.
    /// </summary>
    public bool FoundInFile { get; set; }
    
    /// <summary>
    /// The original header name of the column as found in the file.
    /// </summary>
    public string OriginalHeaderName { get; set; } = string.Empty;
}