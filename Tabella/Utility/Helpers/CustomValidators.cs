using Tabella.Utility.Providers.Interfaces;
using MessageKit.Utility.Builders;
using Tabella.Utility.Delegates;
using Tabella.Utility.Contexts;
using System.Globalization;
using Tabella.Data.Results;
using Tabella._Common;

namespace Tabella.Utility.Helpers;

/// <summary>
/// Custom validation operations class
/// </summary>
public static class CustomValidators
{
    private static ITabellaMessageTemplatesProvider Templates => CustomValidatorContextAccessor.Current.Templates;

    /// <summary>
    /// Validates that the input is a numeric value within the specified range.
    /// </summary>
    /// <param name="min"></param>
    /// <param name="max"></param>
    /// <param name="allowNullValues"></param>
    /// <returns></returns>
    public static CustomValidatorDelegate NumericRangeValidator(double min, double max, bool allowNullValues = false)
    {
        return inputValue =>
        {
            if(inputValue is null)
            {
                return allowNullValues
                    ? CustomValidationResult.Success()
                    : CustomValidationResult.Failure(new MessageBuilder(Templates.ValidationMessageNullNotAllowed));
            }

            return inputValue switch
            {
                int intValue when intValue >= min && intValue <= max => CustomValidationResult.Success(),
                double doubleValue when doubleValue >= min && doubleValue <= max => CustomValidationResult.Success(),
                float floatValue when floatValue >= min && floatValue <= max => CustomValidationResult.Success(),
                decimal decimalValue when(double)decimalValue >= min && (double)decimalValue <= max => CustomValidationResult.Success(),
                string stringValue when double.TryParse(stringValue, out double parsed) && parsed >= min && parsed <= max => CustomValidationResult.Success(),
                _ => CustomValidationResult.Failure(new MessageBuilder(Templates.ValidationMessageInvalidNumericRange)
                    .With(TabellaConstants.MessagePlaceholderCurrentValue, inputValue.ToString() ?? string.Empty)
                    .With("Min", min.ToString(CultureInfo.InvariantCulture))
                    .With("Max", max.ToString(CultureInfo.InvariantCulture))
                )
            };
        };
    }
    
    /// <summary>
    /// Validates that the input string's length is within the specified range.
    /// </summary>
    /// <param name="minLength"></param>
    /// <param name="maxLength"></param>
    /// <param name="allowNullValues"></param>
    /// <returns></returns>
    public static CustomValidatorDelegate StringLengthValidator(int? minLength, int? maxLength, bool allowNullValues = false)
    {
        return inputValue =>
        {
            if(!minLength.HasValue && maxLength.HasValue) return CustomValidationResult.Success();
            
            if(inputValue is null)
            {
                return allowNullValues
                    ? CustomValidationResult.Success()
                    : CustomValidationResult.Failure(new MessageBuilder(Templates.ValidationMessageNullNotAllowed));
            }

            if(inputValue is not string stringValue)
            {
                return CustomValidationResult.Failure(new MessageBuilder(Templates.ValidationMessageInvalidStringType)
                    .With(TabellaConstants.MessagePlaceholderCurrentValue, inputValue.ToString() ?? string.Empty));
            }

            bool minLengthNotReached = minLength.HasValue && stringValue.Length < minLength;
            bool maxLengthExceeded = maxLength.HasValue && stringValue.Length > maxLength.Value;

            if (minLengthNotReached || maxLengthExceeded)
            {
                return CustomValidationResult.Failure(new MessageBuilder(Templates.ValidationMessageInvalidStringLength)
                    .With(TabellaConstants.MessagePlaceholderCurrentValue, inputValue.ToString() ?? string.Empty)
                    .With("MinLength", minLength.ToString() ?? TabellaConstants.MessageRangeNoEndSymbol)
                    .With("MaxLength", maxLength?.ToString() ?? TabellaConstants.MessageRangeNoEndSymbol)
                );
            }

            return CustomValidationResult.Success();
        };
    }

    /// <summary>
    /// Validates that the input decimal value conforms to the specified precision and scale.
    /// </summary>
    /// <param name="precision"></param>
    /// <param name="scale"></param>
    /// <param name="allowNullValues"></param>
    /// <returns></returns>
    public static CustomValidatorDelegate DecimalPrecisionScaleValidator(int precision, int scale, bool allowNullValues = false)
    {
        return inputValue =>
        {
            if(inputValue is null)
            {
                return allowNullValues
                    ? CustomValidationResult.Success()
                    : CustomValidationResult.Failure(new MessageBuilder(Templates.ValidationMessageNullNotAllowed));
            }

            if(inputValue is not decimal decimalValue)
            {
                if(inputValue is string stringValue && decimal.TryParse(stringValue, out decimal parsedDecimal))
                {
                    decimalValue = parsedDecimal;
                }
                else
                {
                    return CustomValidationResult.Failure(new MessageBuilder(Templates.ValidationMessageInvalidDecimalType)
                        .With(TabellaConstants.MessagePlaceholderCurrentValue, inputValue.ToString() ?? string.Empty));
                }
            }

            decimalValue = Math.Abs(decimalValue);
            string[] parts = decimalValue.ToString(CultureInfo.InvariantCulture).Split('.');

            int integerDigits = parts[0].TrimStart('0').Length;
            int fractionalDigits = parts.Length > 1 ? parts[1].TrimEnd('0').Length : 0;

            if(integerDigits + fractionalDigits > precision || fractionalDigits > scale)
            {
                return CustomValidationResult.Failure(new MessageBuilder(Templates.ValidationMessageInvalidDecimalPrecisionScale)
                        .With(TabellaConstants.MessagePlaceholderCurrentValue, inputValue.ToString() ?? string.Empty)
                        .With("Precision", precision.ToString())
                        .With("Scale", scale.ToString())
                );
            }

            return CustomValidationResult.Success();
        };
    }
}