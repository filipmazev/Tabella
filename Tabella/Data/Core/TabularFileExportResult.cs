namespace Tabella.Data.Core;

/// <summary>
/// Represents the result of a tabular file export operation, including success status and any warnings.
/// </summary>
public class TabularFileExportResult
{
    /// <summary>
    /// Indicates whether the export operation was successful.
    /// </summary>
    public required bool IsSuccess { get; set; }
    
    /// <summary>
    /// The list of warnings generated during the export operation.
    /// </summary>
    public required List<string> Warnings { get; set; }
}