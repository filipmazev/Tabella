using System.ComponentModel.DataAnnotations;
using System.Security.Cryptography;
using Tabella.Utility.Attributes;
using System.Reflection;
using System.Text;
using Tabella._Common;

namespace Tabella.Data.Core;

/// <summary>
/// Base class for import models that generate unique keys based on composite properties.
/// </summary>
public class ImportModel
{
    /// <summary>
    /// The generated unique key for the model instance.
    /// </summary>
    public UInt128? Key { get; private set; }
    
    /// <summary>
    /// The generated lookup key for the model instance.
    /// </summary>
    public UInt128? LookupKey { get; private set; }
    
    /// <summary>
    /// The header names used in key generation.
    /// </summary>
    public string? KeyHeaderNames { get; private set; }
    
    /// <summary>
    /// The row index associated with the model instance, if applicable.
    /// </summary>
    public int? RowIndex { get; private set; }
    
    private readonly Dictionary<Type, PropertyInfo[]> _keyPropertiesCache = new();
    private readonly Dictionary<Type, Dictionary<string, PropertyInfo[]>> _lookupKeyPropertiesCache = new();
    private readonly Dictionary<Type, Dictionary<string, PropertyInfo>> _lookupKeyTargetPropertiesCache = new();

    /// <summary>
    /// Generates the unique key and lookup key for the model instance based on composite key properties.
    /// </summary>
    /// <param name="rowIndex"></param>
    public void GenerateKeys(int? rowIndex = null)
    {
        RowIndex = rowIndex;
        
        PropertyInfo[] properties = GetKeyProperties();

        (UInt128 key, string? headerNames) = GenerateKeyWithHeaders(properties);

        Key = key;
        LookupKey = GenerateLookupKey(properties);
        KeyHeaderNames = headerNames;

        Dictionary<string, PropertyInfo[]> lookupKeyGroups = GetLookupKeyProperties();

        foreach ((string targetProperty, PropertyInfo[] keyProperties) in lookupKeyGroups)
        {
            UInt128? generatedKey = GenerateKey(keyProperties);
            SetLookupKeyValue(targetProperty, generatedKey);
        }
    }

    private UInt128 GenerateKey(PropertyInfo[] keyProperties)
    {
        if(keyProperties.Length == 0) return GenerateRandomUInt128();

        string[] concatenatedValue = keyProperties
            .Select(p => FormatValue(p.GetValue(this)))
            .ToArray();

        return GenerateHash(concatenatedValue);
    }

    private UInt128? GenerateLookupKey(PropertyInfo[] keyProperties)
    {
        if(keyProperties.Length == 0) return null;

        string[] concatenatedValue = keyProperties
            .Where(p => p.GetCustomAttribute<CompositeKeyPartAttribute>()?.ExcludeFromLookupKey == false)
            .Select(p => FormatValue(p.GetValue(this)))
            .ToArray();

        return GenerateHash(concatenatedValue);
    }

    private PropertyInfo[] GetKeyProperties()
    {
        Type type = GetType();

        if(_keyPropertiesCache.TryGetValue(type, out PropertyInfo[]? keyProperties)) return keyProperties;

        keyProperties = type.GetProperties()
            .Where(p => p.GetCustomAttribute<CompositeKeyPartAttribute>() is not null &&
                        (p.GetCustomAttribute<RequiredAttribute>() is null ||
                         !string.IsNullOrEmpty(p.GetValue(this)?.ToString())))
            .ToArray();

        _keyPropertiesCache[type] = keyProperties;

        return keyProperties;
    }

    private Dictionary<string, PropertyInfo[]> GetLookupKeyProperties()
    {
        Type type = GetType();

        if (_lookupKeyPropertiesCache.TryGetValue(type, out Dictionary<string, PropertyInfo[]>? lookupKeyGroups)) return lookupKeyGroups;
        Dictionary<string, PropertyInfo> targetProperties = new();

        lookupKeyGroups = type.GetProperties()
            .SelectMany(p => p.GetCustomAttributes<LookupCompositeKeyPartAttribute>()
                .Select(attr => new { attr.TargetPropertyName, Property = p }))
            .GroupBy(x => x.TargetPropertyName)
            .ToDictionary(g => g.Key, g => g.Select(x => x.Property).ToArray());

        foreach(KeyValuePair<string, PropertyInfo[]> group in lookupKeyGroups)
        {
            PropertyInfo? targetProperty = type.GetProperty(group.Key);

            if(targetProperty != null)
            {
                targetProperties[group.Key] = targetProperty;
            }
        }

        _lookupKeyPropertiesCache[type] = lookupKeyGroups;
        _lookupKeyTargetPropertiesCache[type] = targetProperties;
        return lookupKeyGroups;
    }

    private (UInt128 Key, string? HeaderNames) GenerateKeyWithHeaders(PropertyInfo[] keyProperties)
    {
        if (keyProperties.Length == 0) 
            return RowIndex.HasValue 
                ? (GenerateHash([$"{nameof(RowIndex)}:{RowIndex.Value}"]), null)
                : (GenerateRandomUInt128(), null);

        string[] concatenatedValue = new string[keyProperties.Length];
        string[] headerNames = new string[keyProperties.Length];

        for (int i = 0; i < keyProperties.Length; i++)
        {
            PropertyInfo prop = keyProperties[i];
            concatenatedValue[i] = FormatValue(prop.GetValue(this));
            CompositeKeyPartAttribute? attr = prop.GetCustomAttribute<CompositeKeyPartAttribute>();
            headerNames[i] = attr?.HeaderName ?? string.Empty;
        }

        UInt128 key = GenerateHash(concatenatedValue);
        string headersString = string.Join(", ", headerNames.Where(h => !string.IsNullOrEmpty(h)));

        return (key, headersString);
    }
    
    private void SetLookupKeyValue(string targetPropertyName, UInt128? value)
    {
        if(_lookupKeyTargetPropertiesCache.TryGetValue(GetType(), out Dictionary<string, PropertyInfo>? targetProperties) &&
            targetProperties.TryGetValue(targetPropertyName, out PropertyInfo? targetProperty))
        {
            targetProperty.SetValue(this, value);
        }
    }

    private static UInt128 GenerateHash(string[] values)
    {
        byte[] input = Encoding.UTF8.GetBytes(string.Join("|", values));
        byte[] hashBytes = SHA256.HashData(input);

        ulong lower = BitConverter.ToUInt64(hashBytes, 0);
        ulong upper = BitConverter.ToUInt64(hashBytes, 8);

        return (UInt128)upper << 64 | lower;
    }
    
    private static UInt128 GenerateRandomUInt128()
    {
        Span<byte> bytes = stackalloc byte[16]; 
        RandomNumberGenerator.Fill(bytes);

        ulong low = BitConverter.ToUInt64(bytes[..8]);
        ulong high = BitConverter.ToUInt64(bytes[8..]);

        return ((UInt128)high << 64) | low;
    }

    private static string FormatValue(object? value)
    {
        return value switch
        {
            decimal decimalValue => decimalValue.ToString(TabellaConstants.DecimalFormat),
            DateTime dateTimeValue => dateTimeValue.ToString(TabellaConstants.DateFormat),
            Enum enumValue => enumValue.ToString(),
            bool boolValue => boolValue ? TabellaConstants.StringBoolTrue : TabellaConstants.StringBoolFalse,
            _ => value?.ToString()?.Trim() ?? string.Empty
        };
    }

    /// <summary>
    /// Returns a string representation of the model by concatenating all public property values.
    /// </summary>
    /// <returns></returns>
    public override string ToString()
    {
        PropertyInfo[] properties = GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);

        string[] values = properties
            .Select(p => p.GetValue(this)?.ToString() ?? string.Empty)
            .ToArray();

        return string.Join(" ", values);
    }
}