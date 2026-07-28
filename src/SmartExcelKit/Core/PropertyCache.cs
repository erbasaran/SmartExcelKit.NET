using System.Reflection;

namespace SmartExcelKit.Core;

/// <summary>
/// Thread-safe reflection property cache for POCO mapping to avoid repeated reflection overhead.
/// </summary>
internal static class PropertyCache<T> where T : class
{
    /// <summary>
    /// Dictionary of writeable property names (uppercase) to PropertyInfo objects.
    /// </summary>
    public static readonly Dictionary<string, PropertyInfo> WriteableProperties =
        typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                 .Where(p => p.CanWrite)
                 .ToDictionary(p => p.Name.ToUpperInvariant(), p => p);

    /// <summary>
    /// List of readable PropertyInfo objects.
    /// </summary>
    public static readonly List<PropertyInfo> ReadableProperties =
        typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                 .Where(p => p.CanRead)
                 .ToList();
}
