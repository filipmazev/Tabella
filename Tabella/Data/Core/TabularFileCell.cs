namespace Tabella.Data.Core;

/// <summary>
/// Represents a cell in a tabular file with its value and position.
/// </summary>
public class TabularFileCell
{
    /// <summary>
    /// The value contained in the cell.
    /// </summary>
    public required object? CellValue { get; set; }
    
    /// <summary>
    /// The zero-based index of the column where the cell is located.
    /// </summary>
    public required int ColumnIndex { get; set; }
    
    /// <summary>
    /// The zero-based index of the row where the cell is located.
    /// </summary>
    public required int RowIndex { get; set; }
    
    /// <summary>
    /// The name of the sheet containing the cell, if applicable.
    /// </summary>
    public string? SheetName { get; set; }
}