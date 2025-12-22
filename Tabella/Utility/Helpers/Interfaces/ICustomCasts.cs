using Tabella.Data.Results;

namespace Tabella.Utility.Helpers.Interfaces;

/// <summary>
/// Interface for custom casting methods.
/// </summary>
public interface ICustomCasts
{
    /// <summary>
    /// Casts the input to a DateTime object.
    /// </summary>
    /// <param name="input"></param>
    /// <returns></returns>
    CustomCastResult ToDate(object? input);
    
    /// <summary>
    /// Casts the input to a DateTime object with minimum hours (00:00:00).
    /// </summary>
    /// <param name="input"></param>
    /// <returns></returns>
    CustomCastResult ToMinHoursDate(object? input);
    
    
    /// <summary>
    /// Casts the input to a DateTime object with maximum hours (23:59:59).
    /// </summary>
    /// <param name="input"></param>
    /// <returns></returns>
    CustomCastResult ToMaxHoursDate(object? input);
    
    /// <summary>
    /// Casts the input to a boolean value.
    /// </summary>
    /// <param name="input"></param>
    /// <returns></returns>
    CustomCastResult ToBool(object? input);
    
    /// <summary>
    /// Casts the input to an integer value.
    /// </summary>
    /// <param name="input"></param>
    /// <returns></returns>
    
    CustomCastResult ToInt(object? input);
    /// <summary>
    /// Casts the input to a decimal value.
    /// </summary>
    /// <param name="input"></param>
    /// <returns></returns>
    CustomCastResult ToDecimal(object? input);
    
    /// <summary>
    /// Casts the input to a string value.
    /// </summary>
    /// <param name="input"></param>
    /// <returns></returns>
    CustomCastResult ToString(object? input);
    
    /// <summary>
    /// Casts the input to a nullable string value.
    /// </summary>
    /// <param name="input"></param>
    /// <returns></returns>
    CustomCastResult ToNullableString(object? input);
}