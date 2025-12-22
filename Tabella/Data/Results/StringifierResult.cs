namespace Tabella.Data.Results;

/// <summary>
/// Result of a stringification operation.
/// </summary>
public class StringifierResult
{
    /// <summary>
    /// The resulting string value if the operation was successful; otherwise, null.
    /// </summary>
    public string? StringValue { get; private set; }
    
    /// <summary>
    /// Indicates whether the stringification was successful.
    /// </summary>
    public bool IsValid { get; private set; }

    /// <summary>
    /// Creates a successful stringification result.
    /// </summary>
    /// <param name="stringValue"></param>
    /// <returns></returns>
    public static StringifierResult Success(string stringValue) => new() { StringValue = stringValue, IsValid = true };
    
    /// <summary>
    /// Creates a failed stringification result.
    /// </summary>
    /// <returns></returns>
    public static StringifierResult Failure() => new() { IsValid = false };
}