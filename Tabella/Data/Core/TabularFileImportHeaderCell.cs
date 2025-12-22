namespace Tabella.Data.Core;

/// <summary>
/// Represents a header cell in a tabular file import, including its formatted value and associated column mapping.
/// </summary>
public class TabularFileImportHeaderCell : TabularFileCell
{
    /// <summary>
    /// The formatted value of the header cell.
    /// </summary>
    public required string FormatedValue { get; set; }
    
    /// <summary>
    /// The column mapping associated with this header cell.
    /// </summary>
    public required TabularFileColumnMapping ColumnMapping { get; set; }
}