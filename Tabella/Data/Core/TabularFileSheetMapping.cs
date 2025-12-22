namespace Tabella.Data.Core;

/// <summary>
/// Represents the mapping of a sheet in a tabular file to an entity type and its column mappings.
/// </summary>
public class TabularFileSheetMapping
{
    /// <summary>
    /// The entity type associated with this sheet mapping.
    /// </summary>
    public required Type EntityType { get; set; }
    
    /// <summary>
    /// The dictionary of column mappings for this sheet, keyed by column names.
    /// </summary>
    public required Dictionary<string, TabularFileColumnMapping> ColumnMappings { get; set; }
    
    /// <summary>
    /// Indicates whether the sheet was found in the file.
    /// </summary>
    public bool FoundInFile { get; set; }
    
    /// <summary>
    /// The original name of the sheet in the tabular file.
    /// </summary>
    public string OriginalSheetName { get; set; } = string.Empty;
}