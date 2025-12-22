namespace Tabella.Data.Core;

/// <summary>
/// Represents the import model for a tabular file, including sheet name and column data.
/// </summary>
public class TabularFileImportModel : ImportModel
{
    /// <summary>
    /// The name of the sheet from which the data was imported.
    /// </summary>
    public string? SheetName { get; private set; }
    
    /// <summary>
    /// The dictionary containing column data, keyed by column names.
    /// </summary>
    public Dictionary<string, TabularFileColumnData> ColumnData { get; private set; } = new();

    /// <summary>
    /// Sets the cell information based on the provided processed object.
    /// </summary>
    /// <param name="processedObject"></param>
    public void SetCellInfo(TabularFileProcessedObject processedObject)
    {
        SheetName = processedObject.SheetName;
        ColumnData = processedObject.ColumnData;
    }
}