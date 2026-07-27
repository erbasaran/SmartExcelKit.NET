using System;
using System.Globalization;
using SmartExcelKit.Core;
using SmartExcelKit.Styles;

namespace SmartExcelKit;

/// <summary>
/// Represents a cell wrapper providing clean, intuitive access to values, formulas, comments, and styles.
/// </summary>
public sealed class ExcelCell : IEquatable<ExcelCell>
{
    private readonly ExcelWorksheet _worksheet;
    private readonly CellAddress _address;

    /// <summary>
    /// Gets the cell address coordinates.
    /// </summary>
    public CellAddress Address => _address;

    /// <summary>
    /// Gets the 1-based row number of the cell.
    /// </summary>
    public int RowNumber => _address.Row;

    /// <summary>
    /// Gets the 1-based column number of the cell.
    /// </summary>
    public int ColumnNumber => _address.Column;

    /// <summary>
    /// Gets the worksheet containing this cell.
    /// </summary>
    public ExcelWorksheet Worksheet => _worksheet;

    /// <summary>
    /// Gets a value indicating whether the cell has a non-null, non-empty value.
    /// </summary>
    public bool HasValue
    {
        get
        {
            var val = Value;
            if (val == null || val is DBNull) return false;
            if (val is string s) return s.Length > 0;
            return true;
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelCell"/> class.
    /// </summary>
    /// <param name="worksheet">The parent worksheet.</param>
    /// <param name="address">The cell address coordinates.</param>
    internal ExcelCell(ExcelWorksheet worksheet, CellAddress address)
    {
        _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        _address = address;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelCell"/> class using 1-based row and column numbers.
    /// </summary>
    /// <param name="worksheet">The parent worksheet.</param>
    /// <param name="row">The 1-based row index.</param>
    /// <param name="column">The 1-based column index.</param>
    public ExcelCell(ExcelWorksheet worksheet, int row, int column)
        : this(worksheet, new CellAddress(row, column))
    {
    }

    /// <summary>
    /// Gets or sets the cell value.
    /// </summary>
    public object? Value
    {
        get => _worksheet.GetCellValue(_address);
        set => _worksheet.SetCellValue(_address, value);
    }

    /// <summary>
    /// Gets or sets the cell formula (without the leading '=' sign).
    /// </summary>
    public string? Formula
    {
        get => _worksheet.GetCellFormula(_address);
        set => _worksheet.SetCellFormula(_address, value);
    }

    /// <summary>
    /// Gets or sets the cell style.
    /// </summary>
    public ExcelStyle Style
    {
        get => _worksheet.GetCellStyle(_address);
        set => _worksheet.SetCellStyle(_address, value);
    }

    /// <summary>
    /// Gets or sets the comment associated with this cell.
    /// </summary>
    public string? Comment
    {
        get => _worksheet.GetCellComment(_address);
        set => _worksheet.SetCellComment(_address, value);
    }

    /// <summary>
    /// Gets or sets the hyperlink associated with this cell.
    /// </summary>
    public string? Hyperlink
    {
        get => _worksheet.GetCellHyperlink(_address);
        set => _worksheet.SetCellHyperlink(_address, value);
    }

    /// <summary>
    /// Converts and returns the cell value converted to the requested type <typeparamref name="T"/>.
    /// </summary>
    /// <typeparam name="T">The target type.</typeparam>
    /// <returns>The converted value or default if empty or conversion fails.</returns>
    public T? GetValue<T>()
    {
        var val = Value;
        if (val == null || val is DBNull)
            return default;

        if (val is T direct)
            return direct;

        var targetType = Nullable.GetUnderlyingType(typeof(T)) ?? typeof(T);

        try
        {
            if (targetType.IsEnum)
            {
                if (val is string sVal) return (T)Enum.Parse(targetType, sVal, true);
                return (T)Enum.ToObject(targetType, val);
            }

            if (targetType == typeof(Guid))
            {
                return (T)(object)Guid.Parse(val.ToString()!);
            }

            if (targetType == typeof(bool))
            {
                return (T)(object)GetBoolean();
            }

            if (targetType == typeof(DateTime))
            {
                return (T)(object)GetDateTime();
            }

            return (T)Convert.ChangeType(val, targetType, CultureInfo.InvariantCulture);
        }
        catch
        {
            return default;
        }
    }

    /// <summary>
    /// Attempts to convert the cell value to the requested type <typeparamref name="T"/>.
    /// </summary>
    /// <typeparam name="T">The target type.</typeparam>
    /// <param name="value">The output converted value if successful.</param>
    /// <returns>True if conversion succeeded, false otherwise.</returns>
    public bool TryGetValue<T>(out T value)
    {
        var val = Value;
        if (val == null || val is DBNull)
        {
            value = default!;
            return false;
        }

        try
        {
            var res = GetValue<T>();
            if (res != null)
            {
                value = res;
                return true;
            }
        }
        catch
        {
        }

        value = default!;
        return false;
    }

    /// <summary>
    /// Attempts to convert the cell value to <typeparamref name="T"/>, returning <paramref name="defaultValue"/> if cell is empty or conversion fails.
    /// </summary>
    /// <typeparam name="T">The target type.</typeparam>
    /// <param name="defaultValue">The fallback default value.</param>
    /// <returns>The converted value or <paramref name="defaultValue"/>.</returns>
    public T GetValueOrDefault<T>(T defaultValue = default!)
    {
        if (!HasValue) return defaultValue;
        if (TryGetValue<T>(out var val)) return val;
        return defaultValue;
    }

    /// <summary>
    /// Alias for <see cref="GetString"/> returning the string representation of the value.
    /// </summary>
    public string AsString() => GetString();

    /// <summary>
    /// Alias for <see cref="GetInt32"/> converting value to integer.
    /// </summary>
    public int AsInt32() => GetInt32();

    /// <summary>
    /// Alias for <see cref="GetInt64"/> converting value to long.
    /// </summary>
    public long AsInt64() => GetInt64();

    /// <summary>
    /// Alias for <see cref="GetDouble"/> converting value to double.
    /// </summary>
    public double AsDouble() => GetDouble();

    /// <summary>
    /// Alias for <see cref="GetDecimal"/> converting value to decimal.
    /// </summary>
    public decimal AsDecimal() => GetDecimal();

    /// <summary>
    /// Alias for <see cref="GetBoolean"/> converting value to boolean.
    /// </summary>
    public bool AsBoolean() => GetBoolean();

    /// <summary>
    /// Alias for <see cref="GetDateTime"/> converting value to DateTime.
    /// </summary>
    public DateTime AsDateTime() => GetDateTime();

    /// <summary>
    /// Returns the string representation of the value.
    /// </summary>
    public string GetString() => Value?.ToString() ?? string.Empty;

    /// <summary>
    /// Safely attempts to convert the cell value to a 32-bit signed integer.
    /// </summary>
    public int GetInt32()
    {
        var val = Value;
        if (val is int i) return i;
        if (val is double d) return (int)d;
        if (val is float f) return (int)f;
        if (val is long l) return (int)l;
        if (val is decimal dec) return (int)dec;
        if (val is string s && int.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out int parsed)) return parsed;
        return 0;
    }

    /// <summary>
    /// Safely attempts to convert the cell value to a 64-bit signed integer.
    /// </summary>
    public long GetInt64()
    {
        var val = Value;
        if (val is long l) return l;
        if (val is int i) return i;
        if (val is double d) return (long)d;
        if (val is float f) return (long)f;
        if (val is decimal dec) return (long)dec;
        if (val is string s && long.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out long parsed)) return parsed;
        return 0L;
    }

    /// <summary>
    /// Safely attempts to convert the cell value to a double.
    /// </summary>
    public double GetDouble()
    {
        var val = Value;
        if (val is double d) return d;
        if (val is float f) return f;
        if (val is int i) return i;
        if (val is long l) return l;
        if (val is decimal dec) return (double)dec;
        if (val is string s && double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsed)) return parsed;
        return 0.0;
    }

    /// <summary>
    /// Safely attempts to convert the cell value to a decimal.
    /// </summary>
    public decimal GetDecimal()
    {
        var val = Value;
        if (val is decimal dec) return dec;
        if (val is double d) return (decimal)d;
        if (val is float f) return (decimal)f;
        if (val is int i) return i;
        if (val is long l) return l;
        if (val is string s && decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal parsed)) return parsed;
        return 0m;
    }

    /// <summary>
    /// Safely attempts to convert the cell value to a boolean.
    /// </summary>
    public bool GetBoolean()
    {
        var val = Value;
        if (val is bool b) return b;
        if (val is string s && bool.TryParse(s, out bool parsed)) return parsed;
        if (val is double d) return d != 0.0;
        if (val is int i) return i != 0;
        return false;
    }

    /// <summary>
    /// Safely attempts to convert the cell value to a DateTime.
    /// </summary>
    public DateTime GetDateTime()
    {
        var val = Value;
        if (val is DateTime dt) return dt;
        if (val is string s && DateTime.TryParse(s, out DateTime parsed)) return parsed;
        if (val is double d) return DateTime.FromOADate(d);
        return DateTime.MinValue;
    }

    /// <summary>
    /// Resets the cell value, formula, comment, hyperlink, and style.
    /// </summary>
    public void Clear()
    {
        Value = null;
        Formula = null;
        Comment = null;
        Hyperlink = null;
        Style = default;
    }

    /// <inheritdoc />
    public bool Equals(ExcelCell? other)
    {
        if (other is null) return false;
        if (ReferenceEquals(this, other)) return true;
        return _address.Equals(other._address) && ReferenceEquals(_worksheet, other._worksheet);
    }

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is ExcelCell other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => (_address.GetHashCode() * 397) ^ _worksheet.GetHashCode();

    /// <inheritdoc />
    public override string ToString() => $"{_address.Address} = {Value}";
}
