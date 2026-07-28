using System.Data;
using System.Reflection;
using SmartExcelKit.Core;
using SmartExcelKit.Styles;

namespace SmartExcelKit;

/// <summary>
/// Represents an Excel worksheet containing cells, rows, columns, styles, and layout settings.
/// </summary>
public sealed class ExcelWorksheet
{
    private readonly ExcelWorkbook _workbook;
    private string _name;
    private readonly Dictionary<CellAddress, CellData> _cells = [];
    private readonly Dictionary<int, double> _rowHeights = [];
    private readonly Dictionary<int, double> _columnWidths = [];
    private readonly HashSet<int> _hiddenRows = [];
    private readonly HashSet<int> _hiddenColumns = [];
    private readonly List<ExcelRangeAddress> _mergedRanges = [];
    private readonly Dictionary<int, int> _rowOutlineLevels = [];
    private readonly Dictionary<int, int> _columnOutlineLevels = [];

    // Cell counts per row and column for O(1) empty checks
    private readonly Dictionary<int, int> _rowCellCounts = [];
    private readonly Dictionary<int, int> _colCellCounts = [];

    // Cached bounds
    private int _minRow = int.MaxValue;
    private int _maxRow = 0;
    private int _minCol = int.MaxValue;
    private int _maxCol = 0;
    private bool _boundsDirty = false;

    /// <summary>
    /// Gets the workbook associated with this worksheet.
    /// </summary>
    public ExcelWorkbook Workbook => _workbook;

    /// <summary>
    /// Gets or sets the name of the worksheet.
    /// </summary>
    public string Name
    {
        get => _name;
        set
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("Worksheet name cannot be null or empty.", nameof(value));
            if (value.Length > 31)
                throw new ArgumentException("Worksheet name cannot exceed 31 characters.", nameof(value));
            _name = value;
        }
    }

    /// <summary>
    /// Gets or sets whether the worksheet shows gridlines.
    /// </summary>
    public bool ShowGridlines { get; set; } = true;

    /// <summary>
    /// Gets or sets the worksheet zoom percentage (between 10 and 400).
    /// </summary>
    public int Zoom { get; set; } = 100;

    /// <summary>
    /// Gets or sets the tab color in HEX format (e.g., "FF0000").
    /// </summary>
    public string? TabColor { get; set; }

    /// <summary>
    /// Gets or sets the number of frozen rows at the top.
    /// </summary>
    public int FreezeRows { get; set; }

    /// <summary>
    /// Gets or sets the number of frozen columns on the left.
    /// </summary>
    public int FreezeColumns { get; set; }

    /// <summary>
    /// Gets or sets horizontal split position.
    /// </summary>
    public int SplitHorizontal { get; set; }

    /// <summary>
    /// Gets or sets vertical split position.
    /// </summary>
    public int SplitVertical { get; set; }

    /// <summary>
    /// Gets the list of merged cell ranges.
    /// </summary>
    public IReadOnlyList<ExcelRangeAddress> MergedRanges => _mergedRanges;

    /// <summary>
    /// Gets the collection of Excel Tables in this worksheet.
    /// </summary>
    public Tables.ExcelTableCollection Tables { get; }

    /// <summary>
    /// Gets the collection of conditional formatting rules.
    /// </summary>
    public Formatting.ExcelConditionalFormattingCollection ConditionalFormatting { get; }

    /// <summary>
    /// Gets the collection of data validation rules.
    /// </summary>
    public Validation.ExcelDataValidationCollection DataValidations { get; }

    /// <summary>
    /// Gets the collection of worksheet-scoped named ranges.
    /// </summary>
    public Core.ExcelNamedRangeCollection NamedRanges { get; }

    /// <summary>
    /// Gets page print setup settings.
    /// </summary>
    public PageSetup.ExcelPageSetup PageSetup { get; }

    /// <summary>
    /// Gets embedded images in the worksheet.
    /// </summary>
    public List<Drawings.ExcelImage> Images { get; } = [];

    /// <summary>
    /// Gets embedded charts in the worksheet.
    /// </summary>
    public List<Drawings.ExcelChart> Charts { get; } = [];

    /// <summary>
    /// Gets embedded pivot tables in the worksheet.
    /// </summary>
    public List<Drawings.ExcelPivotTable> PivotTables { get; } = [];

    /// <summary>
    /// Gets or sets the AutoFilter range address.
    /// </summary>
    public ExcelRangeAddress? AutoFilterRange { get; set; }

    /// <summary>
    /// Gets the sheet protection password hash (HEX format). Returns null if sheet is not protected.
    /// </summary>
    public string? ProtectionPasswordHash { get; internal set; }

    /// <summary>
    /// Gets a value indicating whether the worksheet is protected.
    /// </summary>
    public bool IsProtected => ProtectionPasswordHash != null;

    /// <summary>
    /// Gets the dictionary of raw cell data (internal use).
    /// </summary>
    internal IReadOnlyDictionary<CellAddress, CellData> RawCells => _cells;

    internal IReadOnlyDictionary<int, double> CustomColumnWidths => _columnWidths;
    internal IReadOnlyDictionary<int, double> CustomRowHeights => _rowHeights;

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelWorksheet"/> class.
    /// </summary>
    internal ExcelWorksheet(ExcelWorkbook workbook, string name)
    {
        _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
        _name = name;

        Tables = new Tables.ExcelTableCollection(this);
        ConditionalFormatting = new Formatting.ExcelConditionalFormattingCollection(this);
        DataValidations = new Validation.ExcelDataValidationCollection(this);
        NamedRanges = new Core.ExcelNamedRangeCollection(name);
        PageSetup = new PageSetup.ExcelPageSetup(this);
    }

    #region Security & Protection

    /// <summary>
    /// Protects the worksheet with a password using the Excel XOR hashing algorithm.
    /// </summary>
    /// <param name="password">The password to protect the sheet with.</param>
    public void Protect(string password)
    {
        if (string.IsNullOrEmpty(password))
        {
            ProtectionPasswordHash = "0000";
            return;
        }

        int hash = 0;
        for (int i = password.Length - 1; i >= 0; i--)
        {
            char c = password[i];
            hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF);
            hash ^= c;
        }
        hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF);
        hash ^= password.Length;
        hash ^= 0xCE4B;

        ProtectionPasswordHash = hash.ToString("X4");
    }

    /// <summary>
    /// Unprotects the worksheet.
    /// </summary>
    public void Unprotect()
    {
        ProtectionPasswordHash = null;
    }

    #endregion

    #region Navigation, Accessors & Indexers

    /// <summary>
    /// Gets an <see cref="ExcelCell"/> by row and column index (1-based).
    /// </summary>
    /// <param name="row">The 1-based row index.</param>
    /// <param name="column">The 1-based column index.</param>
    public ExcelCell this[int row, int column] => Cell(row, column);

    /// <summary>
    /// Gets an <see cref="ExcelRange"/> by range or cell address string (e.g. "A1" or "A1:C10").
    /// </summary>
    /// <param name="rangeOrAddress">The address or range reference string.</param>
    public ExcelRange this[string rangeOrAddress] => Range(rangeOrAddress);

    /// <summary>
    /// Retrieves a cell by its address coordinates.
    /// </summary>
    public ExcelCell Cell(CellAddress address) => new(this, address);

    /// <summary>
    /// Retrieves a cell by its row and column index (1-based).
    /// </summary>
    public ExcelCell Cell(int row, int column) => new(this, new CellAddress(row, column));

    /// <summary>
    /// Retrieves a cell by its string coordinate (e.g. "A1").
    /// </summary>
    public ExcelCell Cell(string address) => Cell(CellAddress.Parse(address));

    /// <summary>
    /// Retrieves a row by its 1-based row index.
    /// </summary>
    /// <param name="index">The 1-based row index.</param>
    /// <returns>An <see cref="ExcelRow"/> instance.</returns>
    public ExcelRow Row(int index) => new(this, index);

    /// <summary>
    /// Retrieves a column by its 1-based column index.
    /// </summary>
    /// <param name="index">The 1-based column index.</param>
    /// <returns>An <see cref="ExcelColumn"/> instance.</returns>
    public ExcelColumn Column(int index) => new(this, index);

    /// <summary>
    /// Retrieves a range by start and end coordinates.
    /// </summary>
    public ExcelRange Range(int startRow, int startColumn, int endRow, int endColumn) =>
        new(this, new ExcelRangeAddress(startRow, startColumn, endRow, endColumn));

    /// <summary>
    /// Retrieves a range by address string (e.g., "A1:B3" or "A1").
    /// </summary>
    public ExcelRange Range(string rangeAddress) => new(this, ExcelRangeAddress.Parse(rangeAddress));

    #endregion

    #region Bounds & Lazy Enumerations

    /// <summary>
    /// Gets the maximum row index currently populated.
    /// </summary>
    public int MaxRow
    {
        get
        {
            EnsureBoundsValid();
            return _maxRow;
        }
    }

    /// <summary>
    /// Gets the maximum column index currently populated.
    /// </summary>
    public int MaxColumn
    {
        get
        {
            EnsureBoundsValid();
            return _maxCol;
        }
    }

    /// <summary>
    /// Gets the total populated row count (alias for <see cref="MaxRow"/>).
    /// </summary>
    public int RowCount => MaxRow;

    /// <summary>
    /// Gets the total populated column count (alias for <see cref="MaxColumn"/>).
    /// </summary>
    public int ColumnCount => MaxColumn;

    /// <summary>
    /// Gets the number of rows spanned by the used range (0 if empty).
    /// </summary>
    public int UsedRowCount => HasUsedCells ? (MaxRow - MinRow + 1) : 0;

    /// <summary>
    /// Gets the number of columns spanned by the used range (0 if empty).
    /// </summary>
    public int UsedColumnCount => HasUsedCells ? (MaxColumn - MinColumn + 1) : 0;

    internal int MinRow
    {
        get
        {
            EnsureBoundsValid();
            return _minRow == int.MaxValue ? 1 : _minRow;
        }
    }

    internal int MinColumn
    {
        get
        {
            EnsureBoundsValid();
            return _minCol == int.MaxValue ? 1 : _minCol;
        }
    }

    private bool HasUsedCells
    {
        get
        {
            EnsureBoundsValid();
            return _maxRow > 0 && _maxCol > 0;
        }
    }

    /// <summary>
    /// Gets an <see cref="ExcelRange"/> covering all populated cells in the worksheet.
    /// Returns range A1:A1 if sheet is empty.
    /// </summary>
    public ExcelRange UsedRange
    {
        get
        {
            EnsureBoundsValid();
            if (!HasUsedCells)
            {
                return Range(1, 1, 1, 1);
            }
            return Range(MinRow, MinColumn, MaxRow, MaxColumn);
        }
    }

    /// <summary>
    /// Returns a lazy enumeration of all rows up to <see cref="MaxRow"/>.
    /// </summary>
    public IEnumerable<ExcelRow> Rows
    {
        get
        {
            int max = Math.Max(1, MaxRow);
            for (int r = 1; r <= max; r++)
            {
                yield return Row(r);
            }
        }
    }

    /// <summary>
    /// Returns a lazy enumeration of all columns up to <see cref="MaxColumn"/>.
    /// </summary>
    public IEnumerable<ExcelColumn> Columns
    {
        get
        {
            int max = Math.Max(1, MaxColumn);
            for (int c = 1; c <= max; c++)
            {
                yield return Column(c);
            }
        }
    }

    /// <summary>
    /// Returns a lazy enumeration of all used (non-empty) cells in the worksheet.
    /// </summary>
    public IEnumerable<ExcelCell> CellsUsed()
    {
        foreach (var kvp in _cells)
        {
            if (kvp.Value != null && !kvp.Value.IsEmpty)
            {
                yield return Cell(kvp.Key);
            }
        }
    }

    /// <summary>
    /// Returns a lazy enumeration of all rows containing used cells or custom row settings.
    /// </summary>
    public IEnumerable<ExcelRow> RowsUsed()
    {
        EnsureBoundsValid();
        if (!HasUsedCells && _rowHeights.Count == 0 && _hiddenRows.Count == 0) yield break;

        int min = Math.Min(MinRow, _rowHeights.Keys.Concat(_hiddenRows).DefaultIfEmpty(int.MaxValue).Min());
        int max = Math.Max(MaxRow, _rowHeights.Keys.Concat(_hiddenRows).DefaultIfEmpty(0).Max());

        if (min > max || max == 0) yield break;

        for (int r = min; r <= max; r++)
        {
            if ((_rowCellCounts.TryGetValue(r, out int cnt) && cnt > 0) || _rowHeights.ContainsKey(r) || _hiddenRows.Contains(r))
            {
                yield return Row(r);
            }
        }
    }

    /// <summary>
    /// Returns a lazy enumeration of all columns containing used cells or custom column settings.
    /// </summary>
    public IEnumerable<ExcelColumn> ColumnsUsed()
    {
        EnsureBoundsValid();
        if (!HasUsedCells && _columnWidths.Count == 0 && _hiddenColumns.Count == 0) yield break;

        int min = Math.Min(MinColumn, _columnWidths.Keys.Concat(_hiddenColumns).DefaultIfEmpty(int.MaxValue).Min());
        int max = Math.Max(MaxColumn, _columnWidths.Keys.Concat(_hiddenColumns).DefaultIfEmpty(0).Max());

        if (min > max || max == 0) yield break;

        for (int c = min; c <= max; c++)
        {
            if ((_colCellCounts.TryGetValue(c, out int cnt) && cnt > 0) || _columnWidths.ContainsKey(c) || _hiddenColumns.Contains(c))
            {
                yield return Column(c);
            }
        }
    }

    /// <summary>
    /// Gets the first used row in the worksheet, or null if empty.
    /// </summary>
    public ExcelRow? FirstRowUsed()
    {
        EnsureBoundsValid();
        return HasUsedCells ? Row(MinRow) : null;
    }

    /// <summary>
    /// Gets the last used row in the worksheet, or null if empty.
    /// </summary>
    public ExcelRow? LastRowUsed()
    {
        EnsureBoundsValid();
        return HasUsedCells ? Row(MaxRow) : null;
    }

    /// <summary>
    /// Gets the first used column in the worksheet, or null if empty.
    /// </summary>
    public ExcelColumn? FirstColumnUsed()
    {
        EnsureBoundsValid();
        return HasUsedCells ? Column(MinColumn) : null;
    }

    /// <summary>
    /// Gets the last used column in the worksheet, or null if empty.
    /// </summary>
    public ExcelColumn? LastColumnUsed()
    {
        EnsureBoundsValid();
        return HasUsedCells ? Column(MaxColumn) : null;
    }

    private void TrackCellAdded(CellAddress address)
    {
        if (address.Row < _minRow) _minRow = address.Row;
        if (address.Row > _maxRow) _maxRow = address.Row;
        if (address.Column < _minCol) _minCol = address.Column;
        if (address.Column > _maxCol) _maxCol = address.Column;

        _rowCellCounts[address.Row] = _rowCellCounts.TryGetValue(address.Row, out int rCnt) ? rCnt + 1 : 1;
        _colCellCounts[address.Column] = _colCellCounts.TryGetValue(address.Column, out int cCnt) ? cCnt + 1 : 1;
    }

    private void TrackCellRemoved(CellAddress address)
    {
        if (_rowCellCounts.TryGetValue(address.Row, out int rCnt))
        {
            if (rCnt <= 1) _rowCellCounts.Remove(address.Row);
            else _rowCellCounts[address.Row] = rCnt - 1;
        }

        if (_colCellCounts.TryGetValue(address.Column, out int cCnt))
        {
            if (cCnt <= 1) _colCellCounts.Remove(address.Column);
            else _colCellCounts[address.Column] = cCnt - 1;
        }

        _boundsDirty = true;
    }

    private void EnsureBoundsValid()
    {
        if (!_boundsDirty) return;

        _minRow = int.MaxValue;
        _maxRow = 0;
        _minCol = int.MaxValue;
        _maxCol = 0;

        foreach (var kvp in _cells)
        {
            if (kvp.Value != null && !kvp.Value.IsEmpty)
            {
                int r = kvp.Key.Row;
                int c = kvp.Key.Column;
                if (r < _minRow) _minRow = r;
                if (r > _maxRow) _maxRow = r;
                if (c < _minCol) _minCol = c;
                if (c > _maxCol) _maxCol = c;
            }
        }

        _boundsDirty = false;
    }

    #endregion

    #region Internal Cell Operations

    private CellData GetOrCreateCellData(CellAddress address)
    {
        if (!_cells.TryGetValue(address, out var data))
        {
            data = new CellData();
            _cells[address] = data;
            TrackCellAdded(address);
        }
        return data;
    }

    internal object? GetCellValue(CellAddress address)
    {
        return _cells.TryGetValue(address, out var data) ? data.Value : null;
    }

    internal void SetCellValue(CellAddress address, object? value)
    {
        if (value == null)
        {
            if (_cells.TryGetValue(address, out var data))
            {
                data.Value = null;
                data.Formula = null;
                if (data.IsEmpty)
                {
                    _cells.Remove(address);
                    TrackCellRemoved(address);
                }
            }
            return;
        }
        GetOrCreateCellData(address).Value = value;
    }

    internal string? GetCellFormula(CellAddress address)
    {
        return _cells.TryGetValue(address, out var data) ? data.Formula : null;
    }

    internal void SetCellFormula(CellAddress address, string? formula)
    {
        if (string.IsNullOrWhiteSpace(formula))
        {
            if (_cells.TryGetValue(address, out var data))
            {
                data.Formula = null;
                if (data.IsEmpty)
                {
                    _cells.Remove(address);
                    TrackCellRemoved(address);
                }
            }
            return;
        }
        GetOrCreateCellData(address).Formula = formula!.TrimStart('=');
    }

    internal ExcelStyle GetCellStyle(CellAddress address)
    {
        if (_cells.TryGetValue(address, out var data))
        {
            return _workbook.StyleRegistry.GetStyle(data.StyleId);
        }
        return default;
    }

    internal void SetCellStyle(CellAddress address, ExcelStyle style)
    {
        uint styleId = _workbook.StyleRegistry.Register(style);
        GetOrCreateCellData(address).StyleId = styleId;
    }

    internal string? GetCellComment(CellAddress address)
    {
        return _cells.TryGetValue(address, out var data) ? data.Comment : null;
    }

    internal void SetCellComment(CellAddress address, string? comment)
    {
        if (string.IsNullOrWhiteSpace(comment))
        {
            if (_cells.TryGetValue(address, out var data))
            {
                data.Comment = null;
                data.CommentObject = null;
                if (data.IsEmpty)
                {
                    _cells.Remove(address);
                    TrackCellRemoved(address);
                }
            }
            return;
        }
        var cellData = GetOrCreateCellData(address);
        cellData.Comment = comment;
        cellData.CommentObject = new Core.ExcelComment(comment!);
    }

    internal Core.ExcelComment? GetCellCommentObject(CellAddress address)
    {
        return _cells.TryGetValue(address, out var data) ? data.CommentObject : null;
    }

    internal void SetCellCommentObject(CellAddress address, Core.ExcelComment? comment)
    {
        if (comment == null)
        {
            if (_cells.TryGetValue(address, out var data))
            {
                data.CommentObject = null;
                data.Comment = null;
                if (data.IsEmpty)
                {
                    _cells.Remove(address);
                    TrackCellRemoved(address);
                }
            }
            return;
        }
        var cellData = GetOrCreateCellData(address);
        cellData.CommentObject = comment;
        cellData.Comment = comment.Text;
    }

    internal string? GetCellHyperlink(CellAddress address)
    {
        return _cells.TryGetValue(address, out var data) ? data.Hyperlink : null;
    }

    internal void SetCellHyperlink(CellAddress address, string? hyperlink)
    {
        if (string.IsNullOrWhiteSpace(hyperlink))
        {
            if (_cells.TryGetValue(address, out var data))
            {
                data.Hyperlink = null;
                data.HyperlinkObject = null;
                if (data.IsEmpty)
                {
                    _cells.Remove(address);
                    TrackCellRemoved(address);
                }
            }
            return;
        }
        var cellData = GetOrCreateCellData(address);
        cellData.Hyperlink = hyperlink;
        cellData.HyperlinkObject = new Core.ExcelHyperlink(hyperlink!);
    }

    internal Core.ExcelHyperlink? GetCellHyperlinkObject(CellAddress address)
    {
        return _cells.TryGetValue(address, out var data) ? data.HyperlinkObject : null;
    }

    internal void SetCellHyperlinkObject(CellAddress address, Core.ExcelHyperlink? hyperlink)
    {
        if (hyperlink == null)
        {
            if (_cells.TryGetValue(address, out var data))
            {
                data.HyperlinkObject = null;
                data.Hyperlink = null;
                if (data.IsEmpty)
                {
                    _cells.Remove(address);
                    TrackCellRemoved(address);
                }
            }
            return;
        }
        var cellData = GetOrCreateCellData(address);
        cellData.HyperlinkObject = hyperlink;
        cellData.Hyperlink = hyperlink.Target;
    }

    internal Core.RichText? GetCellRichText(CellAddress address)
    {
        return _cells.TryGetValue(address, out var data) ? data.RichText : null;
    }

    internal void SetCellRichText(CellAddress address, Core.RichText? richText)
    {
        if (richText == null || richText.Count == 0)
        {
            if (_cells.TryGetValue(address, out var data))
            {
                data.RichText = null;
                if (data.IsEmpty)
                {
                    _cells.Remove(address);
                    TrackCellRemoved(address);
                }
            }
            return;
        }
        var cellData = GetOrCreateCellData(address);
        cellData.RichText = richText;
        cellData.Value = richText.ToString();
    }

    #endregion

    #region Layout, Panes, Grouping & Filtering

    /// <summary>
    /// Splits the worksheet view at the specified pixel offset.
    /// </summary>
    public void SplitPanes(int horizontalPixels, int verticalPixels)
    {
        SplitHorizontal = Math.Max(0, horizontalPixels);
        SplitVertical = Math.Max(0, verticalPixels);
    }

    /// <summary>
    /// Groups a contiguous range of rows into an outline level.
    /// </summary>
    public void GroupRows(int startRow, int endRow)
    {
        if (startRow < 1 || endRow < startRow) throw new ArgumentOutOfRangeException(nameof(startRow));
        for (int r = startRow; r <= endRow; r++)
        {
            _rowOutlineLevels[r] = _rowOutlineLevels.TryGetValue(r, out int lvl) ? lvl + 1 : 1;
        }
    }

    /// <summary>
    /// Ungroups a range of rows in the outline hierarchy.
    /// </summary>
    public void UngroupRows(int startRow, int endRow)
    {
        if (startRow < 1 || endRow < startRow) throw new ArgumentOutOfRangeException(nameof(startRow));
        for (int r = startRow; r <= endRow; r++)
        {
            if (_rowOutlineLevels.TryGetValue(r, out int lvl) && lvl > 1)
                _rowOutlineLevels[r] = lvl - 1;
            else
                _rowOutlineLevels.Remove(r);
        }
    }

    /// <summary>
    /// Groups a contiguous range of columns into an outline level.
    /// </summary>
    public void GroupColumns(int startColumn, int endColumn)
    {
        if (startColumn < 1 || endColumn < startColumn) throw new ArgumentOutOfRangeException(nameof(startColumn));
        for (int c = startColumn; c <= endColumn; c++)
        {
            _columnOutlineLevels[c] = _columnOutlineLevels.TryGetValue(c, out int lvl) ? lvl + 1 : 1;
        }
    }

    /// <summary>
    /// Ungroups a range of columns in the outline hierarchy.
    /// </summary>
    public void UngroupColumns(int startColumn, int endColumn)
    {
        if (startColumn < 1 || endColumn < startColumn) throw new ArgumentOutOfRangeException(nameof(startColumn));
        for (int c = startColumn; c <= endColumn; c++)
        {
            if (_columnOutlineLevels.TryGetValue(c, out int lvl) && lvl > 1)
                _columnOutlineLevels[c] = lvl - 1;
            else
                _columnOutlineLevels.Remove(c);
        }
    }

    /// <summary>
    /// Gets the outline grouping level of a row.
    /// </summary>
    public int GetRowOutlineLevel(int row) => _rowOutlineLevels.TryGetValue(row, out int lvl) ? lvl : 0;

    /// <summary>
    /// Gets the outline grouping level of a column.
    /// </summary>
    public int GetColumnOutlineLevel(int column) => _columnOutlineLevels.TryGetValue(column, out int lvl) ? lvl : 0;

    /// <summary>
    /// Enables AutoFilter on the specified range string or used range.
    /// </summary>
    public void AutoFilter(string? rangeAddress = null)
    {
        if (string.IsNullOrEmpty(rangeAddress))
        {
            AutoFilterRange = UsedRange.Address;
        }
        else
        {
            AutoFilterRange = ExcelRangeAddress.Parse(rangeAddress!);
        }
    }

    /// <summary>
    /// Clears AutoFilter settings from the worksheet.
    /// </summary>
    public void ClearAutoFilter()
    {
        AutoFilterRange = null;
    }

    /// <summary>
    /// Sorts rows within a range by a target column index using stable, culture-aware sorting.
    /// </summary>
    public void Sort(int startRow, int startColumn, int endRow, int endColumn, int sortColumn, bool ascending = true, System.Globalization.CultureInfo? culture = null, IComparer<object?>? customComparer = null)
    {
        if (startRow >= endRow) return;

        culture ??= System.Globalization.CultureInfo.CurrentCulture;

        var rowsData = new List<Tuple<int, List<object?>>>();
        int numCols = endColumn - startColumn + 1;

        for (int r = startRow; r <= endRow; r++)
        {
            var cellValues = new List<object?>(numCols);
            for (int c = startColumn; c <= endColumn; c++)
            {
                cellValues.Add(GetCellValue(new CellAddress(r, c)));
            }
            rowsData.Add(Tuple.Create(r, cellValues));
        }

        int targetColOffset = sortColumn - startColumn;

        rowsData.Sort((r1, r2) =>
        {
            var val1 = targetColOffset >= 0 && targetColOffset < r1.Item2.Count ? r1.Item2[targetColOffset] : null;
            var val2 = targetColOffset >= 0 && targetColOffset < r2.Item2.Count ? r2.Item2[targetColOffset] : null;

            int cmp;
            if (customComparer != null)
            {
                cmp = customComparer.Compare(val1, val2);
            }
            else
            {
                if (val1 == null && val2 == null) cmp = 0;
                else if (val1 == null) cmp = -1;
                else if (val2 == null) cmp = 1;
                else if (val1 is IComparable comp1 && val1.GetType() == val2.GetType())
                {
                    cmp = comp1.CompareTo(val2);
                }
                else
                {
                    cmp = string.Compare(val1.ToString(), val2.ToString(), true, culture);
                }
            }

            return ascending ? cmp : -cmp;
        });

        for (int i = 0; i < rowsData.Count; i++)
        {
            int targetRow = startRow + i;
            var vals = rowsData[i].Item2;
            for (int c = 0; c < vals.Count; c++)
            {
                SetCellValue(new CellAddress(targetRow, startColumn + c), vals[c]);
            }
        }
    }

    #endregion

    #region Layout Settings

    /// <summary>
    /// Sets the width of a specific column.
    /// </summary>
    /// <param name="column">The 1-based column index.</param>
    /// <param name="width">The width value.</param>
    public void SetColumnWidth(int column, double width)
    {
        if (column < 1) throw new ArgumentOutOfRangeException(nameof(column));
        if (width < 0) throw new ArgumentOutOfRangeException(nameof(width));
        _columnWidths[column] = width;
    }

    /// <summary>
    /// Gets the width of a specific column.
    /// </summary>
    /// <param name="column">The 1-based column index.</param>
    /// <returns>The column width.</returns>
    public double GetColumnWidth(int column)
    {
        if (column < 1) throw new ArgumentOutOfRangeException(nameof(column));
        return _columnWidths.TryGetValue(column, out double w) ? w : 8.43;
    }

    /// <summary>
    /// Sets the height of a specific row.
    /// </summary>
    /// <param name="row">The 1-based row index.</param>
    /// <param name="height">The height value.</param>
    public void SetRowHeight(int row, double height)
    {
        if (row < 1) throw new ArgumentOutOfRangeException(nameof(row));
        if (height < 0) throw new ArgumentOutOfRangeException(nameof(height));
        _rowHeights[row] = height;
    }

    /// <summary>
    /// Gets the height of a specific row.
    /// </summary>
    /// <param name="row">The 1-based row index.</param>
    /// <returns>The row height.</returns>
    public double GetRowHeight(int row)
    {
        if (row < 1) throw new ArgumentOutOfRangeException(nameof(row));
        return _rowHeights.TryGetValue(row, out double h) ? h : 15.0;
    }

    /// <summary>
    /// Sets the visibility of a column.
    /// </summary>
    /// <param name="column">The 1-based column index.</param>
    /// <param name="hidden">True to hide the column, false to show it.</param>
    public void SetColumnHidden(int column, bool hidden)
    {
        if (column < 1) throw new ArgumentOutOfRangeException(nameof(column));
        if (hidden) _hiddenColumns.Add(column);
        else _hiddenColumns.Remove(column);
    }

    /// <summary>
    /// Gets whether a column is hidden.
    /// </summary>
    /// <param name="column">The 1-based column index.</param>
    /// <returns>True if hidden, false otherwise.</returns>
    public bool IsColumnHidden(int column) => _hiddenColumns.Contains(column);

    /// <summary>
    /// Sets the visibility of a row.
    /// </summary>
    /// <param name="row">The 1-based row index.</param>
    /// <param name="hidden">True to hide the row, false to show it.</param>
    public void SetRowHidden(int row, bool hidden)
    {
        if (row < 1) throw new ArgumentOutOfRangeException(nameof(row));
        if (hidden) _hiddenRows.Add(row);
        else _hiddenRows.Remove(row);
    }

    /// <summary>
    /// Gets whether a row is hidden.
    /// </summary>
    /// <param name="row">The 1-based row index.</param>
    /// <returns>True if hidden, false otherwise.</returns>
    public bool IsRowHidden(int row) => _hiddenRows.Contains(row);

    /// <summary>
    /// Merges the specified range of cells.
    /// </summary>
    /// <param name="range">The range address to merge.</param>
    public void MergeCells(ExcelRangeAddress range)
    {
        if (!_mergedRanges.Contains(range))
        {
            _mergedRanges.Add(range);
        }
    }

    /// <summary>
    /// Merges the specified range of cells string.
    /// </summary>
    /// <param name="rangeAddress">The range address string (e.g., "A1:D1").</param>
    public void MergeCells(string rangeAddress)
    {
        if (!string.IsNullOrWhiteSpace(rangeAddress))
        {
            MergeCells(ExcelRangeAddress.Parse(rangeAddress));
        }
    }

    /// <summary>
    /// Unmerges the specified range of cells.
    /// </summary>
    /// <param name="range">The range address to unmerge.</param>
    public void UnmergeCells(ExcelRangeAddress range)
    {
        _mergedRanges.Remove(range);
    }

    /// <summary>
    /// Automatically fits the width of a column based on cell text content lengths.
    /// </summary>
    /// <param name="column">The 1-based column index.</param>
    public void AutoFitColumn(int column)
    {
        int maxLen = 0;
        foreach (var cell in _cells)
        {
            if (cell.Key.Column == column && cell.Value.Value != null)
            {
                int len = cell.Value.Value.ToString()?.Length ?? 0;
                if (len > maxLen) maxLen = len;
            }
        }
        double width = Math.Max(8.43, (maxLen + 2) * 1.1);
        SetColumnWidth(column, width);
    }

    #endregion

    #region Bulk & Matrix Operations

    /// <summary>
    /// Bulk sets cell values starting at the specified row and column.
    /// </summary>
    /// <param name="values">The 2D matrix of values.</param>
    /// <param name="startRow">The 1-based starting row index (default 1).</param>
    /// <param name="startColumn">The 1-based starting column index (default 1).</param>
    public void SetValues(object?[,] values, int startRow = 1, int startColumn = 1)
    {
        if (values == null) throw new ArgumentNullException(nameof(values));
        if (startRow < 1 || startColumn < 1) throw new ArgumentOutOfRangeException("Start coordinates must be >= 1.");

        int rows = values.GetLength(0);
        int cols = values.GetLength(1);

        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                int targetRow = startRow + r;
                int targetCol = startColumn + c;
                if (targetRow > 1048576 || targetCol > 16384) continue;
                SetCellValue(new CellAddress(targetRow, targetCol), values[r, c]);
            }
        }
    }

    /// <summary>
    /// Retrieves a 2D matrix of cell values for the specified bounding region.
    /// </summary>
    /// <param name="startRow">The 1-based start row.</param>
    /// <param name="startColumn">The 1-based start column.</param>
    /// <param name="endRow">The 1-based end row.</param>
    /// <param name="endColumn">The 1-based end column.</param>
    /// <returns>A 2D matrix of cell values.</returns>
    public object?[,] GetValues(int startRow, int startColumn, int endRow, int endColumn)
    {
        var range = Range(startRow, startColumn, endRow, endColumn);
        return range.GetValues();
    }

    /// <summary>
    /// Inserts blank rows at the specified index, shifting existing rows down and updating merged ranges.
    /// </summary>
    /// <param name="startRow">The 1-based starting row index to insert at.</param>
    /// <param name="count">The number of rows to insert (default 1).</param>
    public void InsertRows(int startRow, int count = 1)
    {
        if (startRow < 1) throw new ArgumentOutOfRangeException(nameof(startRow));
        if (count < 1) return;

        var entriesToShift = _cells.Where(kvp => kvp.Key.Row >= startRow).ToList();
        foreach (var entry in entriesToShift)
        {
            _cells.Remove(entry.Key);
        }
        foreach (var entry in entriesToShift)
        {
            var newAddr = new CellAddress(entry.Key.Row + count, entry.Key.Column);
            _cells[newAddr] = entry.Value;
        }

        var heightsToShift = _rowHeights.Where(kvp => kvp.Key >= startRow).ToList();
        foreach (var kvp in heightsToShift)
        {
            _rowHeights.Remove(kvp.Key);
        }
        foreach (var kvp in heightsToShift)
        {
            _rowHeights[kvp.Key + count] = kvp.Value;
        }

        var hRows = _hiddenRows.Where(r => r >= startRow).ToList();
        foreach (int r in hRows) _hiddenRows.Remove(r);
        foreach (int r in hRows) _hiddenRows.Add(r + count);

        // Shift merged ranges
        for (int i = 0; i < _mergedRanges.Count; i++)
        {
            var m = _mergedRanges[i];
            if (m.StartRow >= startRow)
            {
                _mergedRanges[i] = new ExcelRangeAddress(m.StartRow + count, m.StartColumn, m.EndRow + count, m.EndColumn);
            }
            else if (m.EndRow >= startRow)
            {
                _mergedRanges[i] = new ExcelRangeAddress(m.StartRow, m.StartColumn, m.EndRow + count, m.EndColumn);
            }
        }

        _boundsDirty = true;
    }

    /// <summary>
    /// Deletes rows at the specified index, shifting remaining lower rows up and updating merged ranges.
    /// </summary>
    /// <param name="startRow">The 1-based starting row index to delete.</param>
    /// <param name="count">The number of rows to delete (default 1).</param>
    public void DeleteRows(int startRow, int count = 1)
    {
        if (startRow < 1) throw new ArgumentOutOfRangeException(nameof(startRow));
        if (count < 1) return;

        int deleteEnd = startRow + count - 1;

        var toRemove = _cells.Where(kvp => kvp.Key.Row >= startRow && kvp.Key.Row <= deleteEnd).Select(kvp => kvp.Key).ToList();
        foreach (var key in toRemove) _cells.Remove(key);

        var entriesToShift = _cells.Where(kvp => kvp.Key.Row > deleteEnd).ToList();
        foreach (var entry in entriesToShift) _cells.Remove(entry.Key);
        foreach (var entry in entriesToShift)
        {
            var newAddr = new CellAddress(entry.Key.Row - count, entry.Key.Column);
            _cells[newAddr] = entry.Value;
        }

        for (int r = startRow; r <= deleteEnd; r++)
        {
            _rowHeights.Remove(r);
            _hiddenRows.Remove(r);
        }

        var heightsToShift = _rowHeights.Where(kvp => kvp.Key > deleteEnd).ToList();
        foreach (var kvp in heightsToShift) _rowHeights.Remove(kvp.Key);
        foreach (var kvp in heightsToShift) _rowHeights[kvp.Key - count] = kvp.Value;

        var hRows = _hiddenRows.Where(r => r > deleteEnd).ToList();
        foreach (int r in hRows) _hiddenRows.Remove(r);
        foreach (int r in hRows) _hiddenRows.Add(r - count);

        // Shift merged ranges
        for (int i = _mergedRanges.Count - 1; i >= 0; i--)
        {
            var m = _mergedRanges[i];
            if (m.StartRow >= startRow && m.EndRow <= deleteEnd)
            {
                _mergedRanges.RemoveAt(i);
            }
            else if (m.StartRow > deleteEnd)
            {
                _mergedRanges[i] = new ExcelRangeAddress(m.StartRow - count, m.StartColumn, m.EndRow - count, m.EndColumn);
            }
        }

        _boundsDirty = true;
    }

    /// <summary>
    /// Inserts blank columns at the specified index, shifting existing columns right and updating merged ranges.
    /// </summary>
    /// <param name="startColumn">The 1-based starting column index to insert at.</param>
    /// <param name="count">The number of columns to insert (default 1).</param>
    public void InsertColumns(int startColumn, int count = 1)
    {
        if (startColumn < 1) throw new ArgumentOutOfRangeException(nameof(startColumn));
        if (count < 1) return;

        var entriesToShift = _cells.Where(kvp => kvp.Key.Column >= startColumn).ToList();
        foreach (var entry in entriesToShift) _cells.Remove(entry.Key);
        foreach (var entry in entriesToShift)
        {
            var newAddr = new CellAddress(entry.Key.Row, entry.Key.Column + count);
            _cells[newAddr] = entry.Value;
        }

        var widthsToShift = _columnWidths.Where(kvp => kvp.Key >= startColumn).ToList();
        foreach (var kvp in widthsToShift) _columnWidths.Remove(kvp.Key);
        foreach (var kvp in widthsToShift) _columnWidths[kvp.Key + count] = kvp.Value;

        var hCols = _hiddenColumns.Where(c => c >= startColumn).ToList();
        foreach (int c in hCols) _hiddenColumns.Remove(c);
        foreach (int c in hCols) _hiddenColumns.Add(c + count);

        // Shift merged ranges
        for (int i = 0; i < _mergedRanges.Count; i++)
        {
            var m = _mergedRanges[i];
            if (m.StartColumn >= startColumn)
            {
                _mergedRanges[i] = new ExcelRangeAddress(m.StartRow, m.StartColumn + count, m.EndRow, m.EndColumn + count);
            }
            else if (m.EndColumn >= startColumn)
            {
                _mergedRanges[i] = new ExcelRangeAddress(m.StartRow, m.StartColumn, m.EndRow, m.EndColumn + count);
            }
        }

        _boundsDirty = true;
    }

    /// <summary>
    /// Deletes columns at the specified index, shifting remaining rightward columns left and updating merged ranges.
    /// </summary>
    /// <param name="startColumn">The 1-based starting column index to delete.</param>
    /// <param name="count">The number of columns to delete (default 1).</param>
    public void DeleteColumns(int startColumn, int count = 1)
    {
        if (startColumn < 1) throw new ArgumentOutOfRangeException(nameof(startColumn));
        if (count < 1) return;

        int deleteEnd = startColumn + count - 1;

        var toRemove = _cells.Where(kvp => kvp.Key.Column >= startColumn && kvp.Key.Column <= deleteEnd).Select(kvp => kvp.Key).ToList();
        foreach (var key in toRemove) _cells.Remove(key);

        var entriesToShift = _cells.Where(kvp => kvp.Key.Column > deleteEnd).ToList();
        foreach (var entry in entriesToShift) _cells.Remove(entry.Key);
        foreach (var entry in entriesToShift)
        {
            var newAddr = new CellAddress(entry.Key.Row, entry.Key.Column - count);
            _cells[newAddr] = entry.Value;
        }

        for (int c = startColumn; c <= deleteEnd; c++)
        {
            _columnWidths.Remove(c);
            _hiddenColumns.Remove(c);
        }

        var widthsToShift = _columnWidths.Where(kvp => kvp.Key > deleteEnd).ToList();
        foreach (var kvp in widthsToShift) _columnWidths.Remove(kvp.Key);
        foreach (var kvp in widthsToShift) _columnWidths[kvp.Key - count] = kvp.Value;

        var hCols = _hiddenColumns.Where(c => c > deleteEnd).ToList();
        foreach (int c in hCols) _hiddenColumns.Remove(c);
        foreach (int c in hCols) _hiddenColumns.Add(c - count);

        // Shift merged ranges
        for (int i = _mergedRanges.Count - 1; i >= 0; i--)
        {
            var m = _mergedRanges[i];
            if (m.StartColumn >= startColumn && m.EndColumn <= deleteEnd)
            {
                _mergedRanges.RemoveAt(i);
            }
            else if (m.StartColumn > deleteEnd)
            {
                _mergedRanges[i] = new ExcelRangeAddress(m.StartRow, m.StartColumn - count, m.EndRow, m.EndColumn - count);
            }
        }

        _boundsDirty = true;
    }

    #endregion

    #region Data Import & Export

    /// <summary>
    /// Imports a strongly typed collection of objects into the worksheet starting at the specified row and column.
    /// </summary>
    public void Import<T>(IEnumerable<T> collection, int startRow = 1, int startColumn = 1) where T : class
    {
        if (collection == null) throw new ArgumentNullException(nameof(collection));
        if (startRow < 1 || startColumn < 1) throw new ArgumentOutOfRangeException("Start coordinates must be >= 1.");

        var properties = PropertyCache<T>.ReadableProperties;

        // Write Headers
        for (int i = 0; i < properties.Count; i++)
        {
            Cell(startRow, startColumn + i).Value = properties[i].Name;
            Cell(startRow, startColumn + i).Style = new ExcelStyle(font: new ExcelFont(bold: true));
        }

        // Write Rows
        int currentRow = startRow + 1;
        foreach (var item in collection)
        {
            if (item == null) continue;
            for (int i = 0; i < properties.Count; i++)
            {
                var val = properties[i].GetValue(item);
                Cell(currentRow, startColumn + i).Value = val;
            }
            currentRow++;
        }
    }

    /// <summary>
    /// Exports the rows of the worksheet back to a strongly typed collection of objects.
    /// </summary>
    /// <param name="startRow">The 1-based start row index.</param>
    /// <param name="startColumn">The 1-based start column index.</param>
    /// <param name="endRow">Optional 1-based end row index (0 or less for MaxRow).</param>
    public IEnumerable<T> Export<T>(int startRow = 1, int startColumn = 1, int endRow = 0) where T : class, new()
    {
        if (startRow < 1 || startColumn < 1) throw new ArgumentOutOfRangeException("Start coordinates must be >= 1.");

        var properties = PropertyCache<T>.WriteableProperties;

        int limitRow = endRow > 0 ? Math.Min(endRow, MaxRow) : MaxRow;
        if (limitRow < startRow) yield break;

        // Read headers
        var colMap = new Dictionary<int, PropertyInfo>();
        int maxCol = MaxColumn;
        for (int col = startColumn; col <= maxCol; col++)
        {
            string header = Cell(startRow, col).GetString().Trim().ToUpperInvariant();
            if (properties.TryGetValue(header, out var prop))
            {
                colMap[col] = prop;
            }
        }

        for (int r = startRow + 1; r <= limitRow; r++)
        {
            var item = new T();
            bool hasAnyData = false;
            foreach (var mapping in colMap)
            {
                var cellVal = Cell(r, mapping.Key).Value;
                if (cellVal != null)
                {
                    hasAnyData = true;
                    var targetType = Nullable.GetUnderlyingType(mapping.Value.PropertyType) ?? mapping.Value.PropertyType;
                    try
                    {
                        var converted = Convert.ChangeType(cellVal, targetType, System.Globalization.CultureInfo.InvariantCulture);
                        mapping.Value.SetValue(item, converted);
                    }
                    catch
                    {
                        if (targetType == typeof(string))
                        {
                            mapping.Value.SetValue(item, cellVal.ToString());
                        }
                    }
                }
            }
            if (hasAnyData)
            {
                yield return item;
            }
        }
    }

    /// <summary>
    /// Imports a DataTable into the worksheet.
    /// </summary>
    public void Import(DataTable dataTable, int startRow = 1, int startColumn = 1, bool includeHeader = true)
    {
        if (dataTable == null) throw new ArgumentNullException(nameof(dataTable));
        if (startRow < 1 || startColumn < 1) throw new ArgumentOutOfRangeException("Start coordinates must be >= 1.");

        int currentRow = startRow;

        if (includeHeader)
        {
            for (int col = 0; col < dataTable.Columns.Count; col++)
            {
                Cell(currentRow, startColumn + col).Value = dataTable.Columns[col].ColumnName;
                Cell(currentRow, startColumn + col).Style = new ExcelStyle(font: new ExcelFont(bold: true));
            }
            currentRow++;
        }

        foreach (DataRow row in dataTable.Rows)
        {
            for (int col = 0; col < dataTable.Columns.Count; col++)
            {
                Cell(currentRow, startColumn + col).Value = row[col] == DBNull.Value ? null : row[col];
            }
            currentRow++;
        }
    }

    /// <summary>
    /// Exports sheet contents back to a DataTable.
    /// </summary>
    public DataTable ExportToDataTable(int startRow = 1, int startColumn = 1, bool hasHeader = true)
    {
        if (startRow < 1 || startColumn < 1) throw new ArgumentOutOfRangeException("Start coordinates must be >= 1.");

        var dt = new DataTable();
        int maxRow = MaxRow;
        int maxCol = MaxColumn;

        if (maxRow < startRow || maxCol < startColumn)
            return dt;

        int dataStartRow = startRow;

        // Establish columns
        if (hasHeader)
        {
            for (int col = startColumn; col <= maxCol; col++)
            {
                string headerName = Cell(startRow, col).GetString();
                if (string.IsNullOrWhiteSpace(headerName))
                {
                    headerName = $"Column_{col}";
                }
                dt.Columns.Add(headerName, typeof(string));
            }
            dataStartRow++;
        }
        else
        {
            for (int col = startColumn; col <= maxCol; col++)
            {
                dt.Columns.Add($"Column_{col}", typeof(string));
            }
        }

        // Fill data rows
        for (int r = dataStartRow; r <= maxRow; r++)
        {
            var rowValues = new string?[dt.Columns.Count];
            bool hasData = false;
            for (int c = 0; c < dt.Columns.Count; c++)
            {
                var val = Cell(r, startColumn + c).GetString();
                rowValues[c] = val;
                if (!string.IsNullOrEmpty(val)) hasData = true;
            }
            if (hasData)
            {
                dt.Rows.Add(rowValues);
            }
        }

        return dt;
    }

    /// <summary>
    /// Formats the used cells of this worksheet into CSV formatted text.
    /// </summary>
    /// <param name="delimiter">The column delimiter character (default ',').</param>
    /// <param name="quoteCharacter">The quote character for escaping (default '"').</param>
    /// <param name="encoding">Optional encoding (unused for string return).</param>
    /// <param name="culture">Optional culture for formatting primitives.</param>
    public string ToCsv(char delimiter = ',', char quoteCharacter = '"', System.Text.Encoding? encoding = null, System.Globalization.CultureInfo? culture = null)
    {
        return UsedRange.ToCsv(delimiter, quoteCharacter, encoding, culture);
    }

    /// <summary>
    /// Serializes the populated worksheet rows into a JSON string.
    /// </summary>
    /// <param name="hasHeader">Whether top row contains headers for JSON property names (default true).</param>
    public string ToJson(bool hasHeader = true)
    {
        return UsedRange.ToJson(hasHeader);
    }

    /// <summary>
    /// Exports sheet contents back to a DataTable (alias for <see cref="ExportToDataTable"/>).
    /// </summary>
    /// <param name="hasHeader">Whether top row contains column headers (default true).</param>
    public DataTable ToDataTable(bool hasHeader = true)
    {
        return ExportToDataTable(1, 1, hasHeader);
    }

    /// <summary>
    /// Exports sheet rows to POCO objects of type <typeparamref name="T"/> (alias for <see cref="Export{T}"/>).
    /// </summary>
    /// <typeparam name="T">The target POCO class type.</typeparam>
    /// <param name="startRow">The 1-based start row (default 1).</param>
    /// <param name="startColumn">The 1-based start column (default 1).</param>
    public IEnumerable<T> ToObjects<T>(int startRow = 1, int startColumn = 1) where T : class, new()
    {
        return Export<T>(startRow, startColumn);
    }

    /// <summary>
    /// Converts sheet rows to a list of header-mapped dictionaries.
    /// </summary>
    /// <param name="startRow">1-based row index to start from (default 1).</param>
    /// <param name="startColumn">1-based column index to start from (default 1).</param>
    /// <param name="hasHeader">Whether top row contains column headers (default true).</param>
    public List<Dictionary<string, object?>> ToDictionaryList(int startRow = 1, int startColumn = 1, bool hasHeader = true)
    {
        if (MaxRow < startRow || MaxColumn < startColumn) return new List<Dictionary<string, object?>>();
        return Range(startRow, startColumn, MaxRow, MaxColumn).ToDictionaryList(hasHeader);
    }

    #endregion
}
