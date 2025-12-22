using MessageKit.Utility.Builders;

namespace Tabella.Data.Results;

/// <summary>
/// Result of a custom casting operation.
/// </summary>
public class CustomCastResult
{
    /// <summary>
    /// The value after casting.
    /// </summary>
    public object? CastedValue { get; private set; }
    
    /// <summary>
    /// Indicates whether the casting was successful.
    /// </summary>
    public bool IsValid { get; private set; }
    
    /// <summary>
    /// Message providing details about the casting result.
    /// </summary>
    public MessageBuilder? Message { get; private set; }

    /// <summary>
    /// Creates a successful casting result.
    /// </summary>
    /// <param name="castedValue"></param>
    /// <returns></returns>
    public static CustomCastResult Success(object? castedValue) => new() { CastedValue = castedValue, IsValid = true };
    
    /// <summary>
    /// Creates a failed casting result.
    /// </summary>
    /// <param name="message"></param>
    /// <returns></returns>
    public static CustomCastResult Failure(MessageBuilder message) => new() { IsValid = false, Message = message };
}