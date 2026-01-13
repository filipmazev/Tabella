namespace Tabella.Utility.Attributes;

/// <summary>
/// Attribute to mark a property as part of a composite lookup key.
/// </summary>
/// <param name="targetPropertyName"></param>
/// <param name="partId">The name of the header (in the original table) associated with this key part.</param>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
public class LookupCompositeKeyPartAttribute(string targetPropertyName, string partId) : Attribute
{
    /// <summary>
    /// The identifier for the key part.
    /// </summary>
    public string PartId { get; } = partId;

    /// <summary>
    /// The name of the target property associated with this lookup key part.
    /// </summary>
    public string TargetPropertyName { get; } = targetPropertyName;
}