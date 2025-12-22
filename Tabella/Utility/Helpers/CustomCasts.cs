using Tabella.Utility.Providers.Interfaces;
using Tabella.Utility.Helpers.Interfaces;
using MessageKit.Utility.Builders;
using Tabella.Data.Results;
using System.Globalization;
using Tabella._Common;

namespace Tabella.Utility.Helpers;

/// <summary>
/// Custom casting operations class
/// </summary>
/// <param name="templates"></param>
public class CustomCasts(
    ITabellaMessageTemplatesProvider templates
    ) : ICustomCasts
{
    /// <summary>
    /// Casts the input to a DateTime object.
    /// </summary>
    /// <param name="input"></param>
    /// <returns></returns>
    public CustomCastResult ToDate(object? input)
    {
        if(input is DateTime dateTimeValue) return CustomCastResult.Success(dateTimeValue);

        if(input is not string stringValue)
        {
            return CustomCastResult.Failure(new MessageBuilder(templates.CastMessageDate)
                .With(TabellaConstants.MessagePlaceholderCastedValue, input?.ToString() ?? string.Empty));
        }

        if (DateTime.TryParse(stringValue, out DateTime parsedValue))
        {
            return CustomCastResult.Success(DateTime.SpecifyKind(parsedValue, DateTimeKind.Utc));
        }

        if(string.IsNullOrEmpty(stringValue)
           || stringValue.Equals(TabellaConstants.EmptyField)
           || CustomFormatters.ToCleanString(stringValue).Equals(TabellaConstants.EmptyField))
        {
            return CustomCastResult.Success(null);
        }

        return CustomCastResult.Failure(new MessageBuilder(templates.CastMessageDate)
            .With(TabellaConstants.MessagePlaceholderCastedValue, input.ToString() ?? string.Empty));
    }

    /// <summary>
    /// Casts the input to a DateTime object with minimum hours (00:00:00).
    /// </summary>
    /// <param name="input"></param>
    /// <returns></returns>
    public CustomCastResult ToMinHoursDate(object? input)
    {
        CustomCastResult toDateCast = ToDate(input);
        if(!toDateCast.IsValid || toDateCast.CastedValue is not DateTime dateValue) return toDateCast;

        return CustomCastResult.Success(new DateTime(dateValue.Year, dateValue.Month, dateValue.Day, 00, 00, 00));
    }

    /// <summary>
    /// Casts the input to a DateTime object with maximum hours (23:59:59).
    /// </summary>
    /// <param name="input"></param>
    /// <returns></returns>
    public CustomCastResult ToMaxHoursDate(object? input)
    {
        CustomCastResult toDateCast = ToDate(input);
        if(!toDateCast.IsValid || toDateCast.CastedValue is not DateTime dateValue) return toDateCast;

        return CustomCastResult.Success(new DateTime(dateValue.Year, dateValue.Month, dateValue.Day, 23, 59, 59));
    }

    /// <summary>
    /// Casts the input to a boolean value.
    /// </summary>
    /// <param name="input"></param>
    /// <returns></returns>
    public CustomCastResult ToBool(object? input)
    {
        return input switch
        {
            bool boolValue => CustomCastResult.Success(boolValue),
            int or float or double or decimal => CustomCastResult.Success((int)input != 0),
            string stringValue when bool.TryParse(stringValue, out bool stringToBoolValue) => CustomCastResult.Success(
                stringToBoolValue),
            string stringValue when int.TryParse(stringValue, out int intValue) => CustomCastResult.Success(intValue != 0),
            string stringValue when float.TryParse(stringValue, CultureInfo.InvariantCulture, out float floatValue) =>
                CustomCastResult.Success((int)floatValue != 0),
            string stringValue when double.TryParse(stringValue, CultureInfo.InvariantCulture, out double doubleValue)
                => CustomCastResult.Success((int)doubleValue != 0),
            string stringValue when decimal.TryParse(stringValue, CultureInfo.InvariantCulture,
                out decimal decimalValue) => CustomCastResult.Success((int)decimalValue != 0),
            string stringValue => CustomFormatters.ToCleanString(stringValue) switch
            {
                "true" => CustomCastResult.Success(true),
                "false" => CustomCastResult.Success(false),
                _ => CustomCastResult.Failure(new MessageBuilder(templates.CastMessageBool)
                    .With(TabellaConstants.MessagePlaceholderCastedValue, input.ToString() ?? string.Empty))
            },
            _ => CustomCastResult.Failure(new MessageBuilder(templates.CastMessageBool)
                    .With(TabellaConstants.MessagePlaceholderCastedValue, input?.ToString() ?? string.Empty))
        };
    }

    /// <summary>
    /// Casts the input to an integer value.
    /// </summary>
    /// <param name="input"></param>
    /// <returns></returns>
    public CustomCastResult ToInt(object? input)
    {
        return input switch
        {
            int intValue => CustomCastResult.Success(intValue),
            double doubleValue => CustomCastResult.Success((int)doubleValue),
            float floatValue => CustomCastResult.Success((int)floatValue),
            decimal decimalValue => CustomCastResult.Success((int)decimalValue),
            string stringValue when int.TryParse(CustomFormatters.ToNumericString(stringValue), out int parsedValue) => CustomCastResult.Success(
                parsedValue),
            string stringValue when string.IsNullOrEmpty(stringValue) 
                                    || CustomFormatters.ToCleanString(stringValue).Equals(TabellaConstants.EmptyField)
                                    || CustomFormatters.ToCleanString(stringValue).Equals(TabellaConstants.NotAssigned) =>
                CustomCastResult.Success(null),
            _ => CustomCastResult.Failure(new MessageBuilder(templates.CastMessageGeneric)
                    .With(TabellaConstants.MessagePlaceholderCastedValue, input?.ToString() ?? string.Empty))
        };
    }

    /// <summary>
    /// Casts the input to a decimal value.
    /// </summary>
    /// <param name="input"></param>
    /// <returns></returns>
    public CustomCastResult ToDecimal(object? input)
    {
        return input switch
        {
            decimal decimalValue => CustomCastResult.Success(decimalValue),
            int intValue => CustomCastResult.Success((decimal)intValue),
            double doubleValue => CustomCastResult.Success((decimal)doubleValue),
            float floatValue => CustomCastResult.Success((decimal)floatValue),
            string stringValue when decimal.TryParse(CustomFormatters.ToNumericString(stringValue, allowDecimals: true), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal parsedValue)
                => CustomCastResult.Success(parsedValue), 
            string stringValue when string.IsNullOrEmpty(stringValue) 
                                    || CustomFormatters.ToCleanString(stringValue).Equals(TabellaConstants.EmptyField)
                                    || CustomFormatters.ToCleanString(stringValue).Equals(TabellaConstants.NotAssigned) =>
                CustomCastResult.Success(null),
            _ => CustomCastResult.Failure(new MessageBuilder(templates.CastMessageGeneric)
                    .With(TabellaConstants.MessagePlaceholderCastedValue, input?.ToString() ?? string.Empty))
        };
    }
    
    /// <summary>
    /// Casts the input to a string value.
    /// </summary>
    /// <param name="input"></param>
    /// <returns></returns>
    public CustomCastResult ToString(object? input)
    {
        if(input is not string stringValue) return CustomCastResult.Success(input?.ToString() ?? string.Empty);

        bool empty = input.Equals(TabellaConstants.EmptyField);

        return CustomCastResult.Success(empty ? string.Empty : stringValue.Trim());
    }

    /// <summary>
    /// Casts the input to a string value.
    /// </summary>
    /// <param name="input"></param>
    /// <returns></returns>
    public CustomCastResult ToNullableString(object? input)
    {
        if(input is not string stringValue) return CustomCastResult.Success(input?.ToString());

        bool empty = input.Equals(TabellaConstants.EmptyField)
                     || CustomFormatters.ToCleanString(stringValue).Equals(TabellaConstants.NotAssigned);

        return CustomCastResult.Success(empty ? null : stringValue.Trim());
    }
}