using MessageKit.Utility.Builders;

namespace Tabella.Data.Results;

/// <summary>
/// Result of a custom validation operation.
/// </summary>
public class CustomValidationResult
{
    /// <summary>
    /// Indicates whether the validation was successful.
    /// </summary>
    public bool IsValid { get; private set; }
    
    /// <summary>
    /// Message providing details about the validation result.
    /// </summary>
    public MessageBuilder? Message { get; private set; }

    /// <summary>
    /// Creates a successful validation result.
    /// </summary>
    /// <returns></returns>
    public static CustomValidationResult Success() => new() { IsValid = true };
    
    /// <summary>
    /// Creates a failed validation result.
    /// </summary>
    /// <param name="message"></param>
    /// <returns></returns>
    public static CustomValidationResult Failure(MessageBuilder message) => new() { IsValid = false, Message = message };
}