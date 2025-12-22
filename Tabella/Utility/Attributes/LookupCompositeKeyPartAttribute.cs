namespace Tabella.Utility.Attributes;

/// <summary>
/// Attribute to mark a property as part of a composite lookup key.
/// </summary>
/// <param name="targetPropertyName"></param>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
public class LookupCompositeKeyPartAttribute(string targetPropertyName) : Attribute
{
    /// <summary>
    /// The name of the target property associated with this lookup key part.
    /// </summary>
    public string TargetPropertyName { get; } = targetPropertyName;
}