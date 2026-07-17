using SmartExcelKit.Core;
using SmartExcelKit.Styles;
using System.Collections;

namespace SmartExcelKit;

/// <summary>
/// Represents a range of cells, allowing bulk updates, styles, formatting, merging, and iterating.
/// </summary>
public sealed class ExcelRange : IEnumerable<ExcelCell>
{
    private readonly ExcelWorksheet _worksheet;
    private readonly ExcelRangeAddress _address;

    /// <summary>
    /// Gets the range bounding address coordinates.
    /// </summary>
    public ExcelRangeAddress Address => _address;

    /// <summary>
    /// Gets the worksheet containing this range.
    /// </summary>
    public ExcelWorksheet Worksheet => _worksheet;

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelRange"/> class.
    /// </summary>
    internal ExcelRange(ExcelWorksheet worksheet, ExcelRangeAddress address)
    {
        _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        _address = address;
    }

    /// <summary>
    /// Sets the value for all cells within this range.
    /// </summary>
    public object? Value
    {
        set
        {
            for (int r = _address.StartRow; r <= _address.EndRow; r++)
            {
                for (int c = _address.StartColumn; c <= _address.EndColumn; c++)
                {
                    _worksheet.SetCellValue(new CellAddress(r, c), value);
                }
            }
        }
    }

    /// <summary>
    /// Sets the formula for all cells within this range.
    /// </summary>
    public string? Formula
    {
        set
        {
            for (int r = _address.StartRow; r <= _address.EndRow; r++)
            {
                for (int c = _address.StartColumn; c <= _address.EndColumn; c++)
                {
                    _worksheet.SetCellFormula(new CellAddress(r, c), value);
                }
            }
        }
    }

    /// <summary>
    /// Sets the style for all cells within this range.
    /// </summary>
    public ExcelStyle Style
    {
        set
        {
            for (int r = _address.StartRow; r <= _address.EndRow; r++)
            {
                for (int c = _address.StartColumn; c <= _address.EndColumn; c++)
                {
                    _worksheet.SetCellStyle(new CellAddress(r, c), value);
                }
            }
        }
    }

    /// <summary>
    /// Merges the cells in this range.
    /// </summary>
    public void Merge()
    {
        _worksheet.MergeCells(_address);
    }

    /// <summary>
    /// Unmerges the cells in this range.
    /// </summary>
    public void Unmerge()
    {
        _worksheet.UnmergeCells(_address);
    }

    /// <summary>
    /// Clears values, formulas, comments, and styles from the cells in this range.
    /// </summary>
    public void Clear()
    {
        ClearContents();
        ClearFormats();
        for (int r = _address.StartRow; r <= _address.EndRow; r++)
        {
            for (int c = _address.StartColumn; c <= _address.EndColumn; c++)
            {
                var addr = new CellAddress(r, c);
                _worksheet.SetCellComment(addr, null);
                _worksheet.SetCellHyperlink(addr, null);
            }
        }
    }

    /// <summary>
    /// Resets formatting/styles of all cells in the range.
    /// </summary>
    public void ClearFormats()
    {
        Style = default;
    }

    /// <summary>
    /// Clears values and formulas of all cells in the range.
    /// </summary>
    public void ClearContents()
    {
        for (int r = _address.StartRow; r <= _address.EndRow; r++)
        {
            for (int c = _address.StartColumn; c <= _address.EndColumn; c++)
            {
                var addr = new CellAddress(r, c);
                _worksheet.SetCellValue(addr, null);
                _worksheet.SetCellFormula(addr, null);
            }
        }
    }

    /// <summary>
    /// Automatically fits column widths to content for columns inside this range.
    /// </summary>
    public void AutoFitColumns()
    {
        for (int c = _address.StartColumn; c <= _address.EndColumn; c++)
        {
            _worksheet.AutoFitColumn(c);
        }
    }

    /// <summary>
    /// Enumerates all cells sequentially, row-by-row, column-by-column.
    /// </summary>
    public IEnumerator<ExcelCell> GetEnumerator()
    {
        for (int r = _address.StartRow; r <= _address.EndRow; r++)
        {
            for (int c = _address.StartColumn; c <= _address.EndColumn; c++)
            {
                yield return _worksheet.Cell(r, c);
            }
        }
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
