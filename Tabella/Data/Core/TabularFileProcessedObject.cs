namespace Tabella.Data.Core;

/// <summary>
/// Represents a processed object from a tabular file, including its associated sheet name, row index, and column data.
/// </summary>
public class TabularFileProcessedObject
{
    /// <summary>
    /// The processed object extracted from the tabular file.
    /// </summary>
    public required object? Object { get; set; }
    
    /// <summary>
    /// The name of the sheet from which the object was processed.
    /// </summary>
    public string? SheetName { get; set; }
    
    /// <summary>
    /// The index of the row from which the object was processed.
    /// </summary>
    public int? RowIndex { get; set; }
    
    /// <summary>
    /// The dictionary containing column data, keyed by column names.
    /// </summary>
    public required Dictionary<string, TabularFileColumnData> ColumnData { get; set; }
}