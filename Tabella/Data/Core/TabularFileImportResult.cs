namespace Tabella.Data.Core;

/// <summary>
/// Represents the result of a tabular file import operation, including processed objects and messages.
/// </summary>
public class TabularFileImportResult
{
    /// <summary>
    /// The dictionary of processed objects, keyed by their type.
    /// </summary>
    public required Dictionary<Type, List<TabularFileProcessedObject>> ProcessedObjects { get; set; }
    
    /// <summary>
    /// The set of messages generated during the import operation.
    /// </summary>
    public required HashSet<TabularFileImportMessage> Messages { get; set; } = [];
}