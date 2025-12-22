namespace Tabella.Utility.Attributes;

/// <summary>
/// Attribute to mark a property as part of a composite key.
/// </summary>
/// <param name="headerName"></param>
/// <param name="excludeFromLookupKey"></param>
[AttributeUsage(AttributeTargets.Property)]
public sealed class CompositeKeyPartAttribute(string headerName, bool excludeFromLookupKey = false) : Attribute
{
    /// <summary>
    /// The name of the header associated with this key part.
    /// </summary>
    public string HeaderName { get; } = headerName;
    
    /// <summary>
    /// Indicates whether this key part should be excluded from lookup keys.
    /// </summary>
    public bool ExcludeFromLookupKey { get; } = excludeFromLookupKey;
}