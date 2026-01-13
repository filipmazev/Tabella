using System.Text.RegularExpressions;
using Tabella._Common;
using System.Text;

namespace Tabella.Utility.Helpers;

/// <summary>
/// Custom formatters for data processing.
/// </summary>
public static class CustomFormatters
{
    /// <summary>
    /// Cleans the input string by removing non-alphanumeric characters,
    /// </summary>
    /// <param name="input"></param>
    /// <returns></returns>
    public static string ToCleanString(object? input)
    {
        if(input is not string stringValue || string.IsNullOrEmpty(stringValue)) return string.Empty;

        string cleanedString = Regex.Replace(stringValue, @"[^\w\d]", string.Empty);

        cleanedString = cleanedString.ToLower();

        byte[] encodedBytes = Encoding.UTF8.GetBytes(cleanedString);
        string encodedString = Encoding.UTF8.GetString(encodedBytes);

        return encodedString;
    }

    /// <summary>
    /// Converts the input to a numeric string, optionally allowing decimal points.
    /// </summary>
    /// <param name="input"></param>
    /// <param name="allowDecimals"></param>
    /// <returns></returns>
    public static string ToNumericString(object? input, bool allowDecimals = false)
    {
        if(input is not string stringValue || string.IsNullOrEmpty(stringValue)) return string.Empty;

        if(!allowDecimals) return Regex.Replace(stringValue, @"[^\d]", string.Empty);
        
        string numericString = Regex.Replace(stringValue, @"[^\d.,]", string.Empty);

        int lastDot = numericString.LastIndexOf('.');
        int lastComma = numericString.LastIndexOf(',');

        int lastSeparator = Math.Max(lastDot, lastComma);

        if(lastSeparator == -1) return Regex.Replace(numericString, @"[^\d]", string.Empty);

        string integerPart = Regex.Replace(numericString[..lastSeparator], @"[^\d]", string.Empty);
        string decimalPart = Regex.Replace(numericString[(lastSeparator + 1)..], @"[^\d]", string.Empty);

        return $"{integerPart}.{decimalPart}";
    }

    /// <summary>
    /// Converts the input to a padded numeric string.
    /// </summary>
    /// <param name="input"></param>
    /// <param name="totalLength"></param>
    /// <param name="paddingChar"></param>
    /// <param name="minLengthOnly"></param>
    /// <returns></returns>
    public static string? ToPaddedNumericString(
        object? input,
        int totalLength,
        char paddingChar,
        bool minLengthOnly = false)
    {
        string? stringValue = input?.ToString()?.Trim();

        switch (stringValue)
        {
            case null:
                return stringValue;
            case TabellaConstants.EmptyField:
            case TabellaConstants.NotAssigned:
                return null;
        }

        string numericOnly = ToNumericString(stringValue);

        if(string.IsNullOrEmpty(numericOnly) || !Regex.IsMatch(numericOnly, @"^\d+$"))
            return null;

        if(minLengthOnly && numericOnly.Length >= totalLength)
            return numericOnly;

        return numericOnly.PadLeft(totalLength, paddingChar);
    }
}