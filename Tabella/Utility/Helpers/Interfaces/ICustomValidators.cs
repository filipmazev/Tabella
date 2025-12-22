using Tabella.Utility.Delegates;

namespace Tabella.Utility.Helpers.Interfaces;

/// <summary>
/// Interface for custom validation methods.
/// </summary>
public interface ICustomValidators
{
    /// <summary>
    /// Numeric range validator.
    /// </summary>
    /// <param name="min"></param>
    /// <param name="max"></param>
    /// <param name="allowNullValues"></param>
    /// <returns></returns>
    CustomValidatorDelegate NumericRangeValidator(double min, double max, bool allowNullValues = false);
    
    /// <summary>
    /// String length validator.
    /// </summary>
    /// <param name="minLength"></param>
    /// <param name="maxLength"></param>
    /// <param name="allowNullValues"></param>
    /// <returns></returns>
    CustomValidatorDelegate StringLengthValidator(int? minLength, int? maxLength, bool allowNullValues = false);
    /// <summary>
    /// Decimal precision and scale validator.
    /// </summary>
    /// <param name="precision"></param>
    /// <param name="scale"></param>
    /// <param name="allowNullValues"></param>
    /// <returns></returns>
    CustomValidatorDelegate DecimalPrecisionScaleValidator(int precision, int scale, bool allowNullValues = false);
}
