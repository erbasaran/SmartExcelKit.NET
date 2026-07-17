using SmartExcelKit.Exceptions;

namespace SmartExcelKit.Formula;

/// <summary>
/// Tracks cell dependencies and detects circular references in the workbook.
/// </summary>
internal sealed class DependencyGraph
{
    // Key: "SheetName!Row,Col" depends on values: Set of "SheetName!Row,Col"
    private readonly Dictionary<string, HashSet<string>> _adjList = [];

    /// <summary>
    /// Registers the dependencies for a cell. Clears any previous dependencies for that cell.
    /// </summary>
    /// <param name="cellKey">The cell identifier (e.g. "Sheet1!A1").</param>
    /// <param name="dependsOnKeys">The list of cells that this cell depends on.</param>
    /// <exception cref="FormulaException">Thrown if a circular reference is detected.</exception>
    public void SetDependencies(string cellKey, IEnumerable<string> dependsOnKeys)
    {
        if (string.IsNullOrEmpty(cellKey)) throw new ArgumentNullException(nameof(cellKey));

        // Backup existing dependencies in case of a rollback due to cycle detection
        HashSet<string>? backup = null;
        if (_adjList.TryGetValue(cellKey, out var existing))
        {
            backup = new HashSet<string>(existing);
        }

        var newDeps = new HashSet<string>(dependsOnKeys);
        _adjList[cellKey] = newDeps;

        // Perform cycle detection using DFS
        var visited = new HashSet<string>();
        var recursionStack = new HashSet<string>();

        foreach (var key in _adjList.Keys)
        {
            if (CheckCircular(key, visited, recursionStack))
            {
                // Rollback dependencies to avoid leaving the graph in an invalid state
                if (backup != null)
                {
                    _adjList[cellKey] = backup;
                }
                else
                {
                    _adjList.Remove(cellKey);
                }

                throw new FormulaException($"Circular reference detected involving cell '{cellKey}'.", "CIRCULAR_REFERENCE");
            }
        }
    }

    /// <summary>
    /// Clears all registered dependencies for a specific cell.
    /// </summary>
    public void ClearCell(string cellKey)
    {
        if (string.IsNullOrEmpty(cellKey)) return;
        _adjList.Remove(cellKey);
    }

    private bool CheckCircular(string u, HashSet<string> visited, HashSet<string> recursionStack)
    {
        if (recursionStack.Contains(u))
            return true; // Found a cycle!

        if (visited.Contains(u))
            return false; // Already verified this node

        visited.Add(u);
        recursionStack.Add(u);

        if (_adjList.TryGetValue(u, out var neighbors))
        {
            foreach (var v in neighbors)
            {
                if (CheckCircular(v, visited, recursionStack))
                    return true;
            }
        }

        recursionStack.Remove(u);
        return false;
    }
}
