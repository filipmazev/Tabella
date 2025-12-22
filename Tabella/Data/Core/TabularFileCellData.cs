namespace Tabella.Data.Core;

/// <summary>
/// Represents data about a cell in a tabular file, including its column name, index, and value.
/// </summary>
public class TabularFileCellData
{
    /// <summary>
    /// The name of the column where the cell is located.
    /// </summary>
    public string? ColumnName { get; set; }
    
    /// <summary>
    /// The zero-based index of the column where the cell is located.
    /// </summary>
    public int? ColumnIndex { get; set; }
    
    /// <summary>
    /// The value contained in the cell.
    /// </summary>
    public string? CellValue { get; set; }
}