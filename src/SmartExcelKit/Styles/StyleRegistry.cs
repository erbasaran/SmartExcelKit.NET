namespace SmartExcelKit.Styles;

/// <summary>
/// Registers and caches unique style definitions, allocating numerical IDs to minimize memory and XML size.
/// </summary>
public sealed class StyleRegistry
{
    private readonly List<ExcelStyle> _styles = [];
    private readonly Dictionary<ExcelStyle, uint> _styleToId = [];

    /// <summary>
    /// Gets the list of registered styles.
    /// </summary>
    public IReadOnlyList<ExcelStyle> RegisteredStyles => _styles;

    /// <summary>
    /// Initializes a new instance of the <see cref="StyleRegistry"/> class with a default General/Calibri style.
    /// </summary>
    public StyleRegistry()
    {
        // Register default style (ID 0)
        Register(default);
    }

    /// <summary>
    /// Registers a style or returns its existing index.
    /// </summary>
    /// <param name="style">The style structure to register.</param>
    /// <returns>A unique 0-based index identifier for the style.</returns>
    public uint Register(ExcelStyle style)
    {
        if (_styleToId.TryGetValue(style, out uint existingId))
        {
            return existingId;
        }

        uint newId = (uint)_styles.Count;
        _styles.Add(style);
        _styleToId.Add(style, newId);
        return newId;
    }

    /// <summary>
    /// Retrieves style details using its index.
    /// </summary>
    /// <param name="index">The 0-based style index.</param>
    /// <returns>The registered <see cref="ExcelStyle"/>.</returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown if the index is out of bounds.</exception>
    public ExcelStyle GetStyle(uint index)
    {
        if (index >= _styles.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(index), $"Style ID {index} is not registered.");
        }
        return _styles[(int)index];
    }
}
