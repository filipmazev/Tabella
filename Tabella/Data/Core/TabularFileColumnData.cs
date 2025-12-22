namespace Tabella.Data.Core;

/// <summary>
/// Represents data about a column in a tabular file, including its target entity property, object property type, and associated cell data.
/// </summary>
public class TabularFileColumnData
{
    /// <summary>
    /// The type of the object property that this column maps to.
    /// </summary>
    public required Type ObjectPropertyType { get; set; }
    
    /// <summary>
    /// The name of the target entity property that this column maps to.
    /// </summary>
    public required string TargetEntityProperty { get; set; }
    
    /// <summary>
    /// The data of the cell associated with this column.
    /// </summary>
    public TabularFileCellData? CellData { get; set; }
}