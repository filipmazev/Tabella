using Tabella.Utility.Delegates;
using Tabella.Data.Results;

namespace Tabella.Utility.Helpers;

/// <summary>
/// Custom validation operations class
/// </summary>
public static class Stringifiers
{
    /// <summary>
    /// Creates a stringifier that converts any object to its string representation.
    /// </summary>
    /// <param name="allowNullValues"></param>
    /// <returns></returns>
    public static StringifierDelegate FromString(bool allowNullValues = false)
    {
        return input =>
        {
            if (input is null)
                return allowNullValues ? StringifierResult.Success(string.Empty) : StringifierResult.Failure();

            return StringifierResult.Success(input.ToString() ?? string.Empty);
        };
    }
}