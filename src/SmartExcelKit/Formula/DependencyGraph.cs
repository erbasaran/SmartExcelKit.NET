using SmartExcelKit.Exceptions;

namespace SmartExcelKit.Formula;

/// <summary>
/// High-performance directed graph for tracking cell formula dependencies, cycle detection, and topological recalculation order.
/// </summary>
internal sealed class DependencyGraph
{
    // Key: Dependent Cell ("SheetName!Row,Col") -> Set of Prerequisite Cells it depends on
    private readonly Dictionary<string, HashSet<string>> _dependsOn = new(StringComparer.OrdinalIgnoreCase);

    // Key: Prerequisite Cell ("SheetName!Row,Col") -> Set of Dependent Cells that rely on it
    private readonly Dictionary<string, HashSet<string>> _dependents = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Registers the dependencies for a cell. Clears any previous dependencies for that cell.
    /// </summary>
    /// <param name="cellKey">The cell identifier (e.g. "Sheet1!A1").</param>
    /// <param name="dependsOnKeys">The list of prerequisite cells that this cell depends on.</param>
    /// <exception cref="FormulaException">Thrown if a circular reference is detected.</exception>
    public void SetDependencies(string cellKey, IEnumerable<string> dependsOnKeys)
    {
        if (string.IsNullOrEmpty(cellKey)) throw new ArgumentNullException(nameof(cellKey));

        // Remove previous inverse mappings
        if (_dependsOn.TryGetValue(cellKey, out var oldDeps))
        {
            foreach (var dep in oldDeps)
            {
                if (_dependents.TryGetValue(dep, out var set))
                {
                    set.Remove(cellKey);
                }
            }
        }

        var newDeps = new HashSet<string>(dependsOnKeys, StringComparer.OrdinalIgnoreCase);
        _dependsOn[cellKey] = newDeps;

        foreach (var dep in newDeps)
        {
            if (!_dependents.TryGetValue(dep, out var set))
            {
                set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                _dependents[dep] = set;
            }
            set.Add(cellKey);
        }

        // Perform cycle detection using DFS
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var recursionStack = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        if (CheckCircular(cellKey, visited, recursionStack))
        {
            ClearCell(cellKey);
            throw new FormulaException($"Circular reference detected involving cell '{cellKey}'.", "CIRCULAR_REFERENCE");
        }
    }

    /// <summary>
    /// Clears all registered dependencies for a specific cell.
    /// </summary>
    public void ClearCell(string cellKey)
    {
        if (string.IsNullOrEmpty(cellKey)) return;

        if (_dependsOn.TryGetValue(cellKey, out var oldDeps))
        {
            foreach (var dep in oldDeps)
            {
                if (_dependents.TryGetValue(dep, out var set))
                {
                    set.Remove(cellKey);
                }
            }
            _dependsOn.Remove(cellKey);
        }
    }

    /// <summary>
    /// Gets the topological calculation order for cells affected by changes to the specified root cell.
    /// </summary>
    public List<string> GetAffectedCalculationOrder(string rootCellKey)
    {
        var result = new List<string>();
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        void Visit(string key)
        {
            if (!visited.Add(key)) return;

            if (_dependents.TryGetValue(key, out var directDependents))
            {
                foreach (var dependent in directDependents)
                {
                    Visit(dependent);
                }
            }

            if (!string.Equals(key, rootCellKey, StringComparison.OrdinalIgnoreCase))
            {
                result.Add(key);
            }
        }

        Visit(rootCellKey);
        result.Reverse(); // Ensure proper prerequisite order
        return result;
    }

    private bool CheckCircular(string u, HashSet<string> visited, HashSet<string> recursionStack)
    {
        if (recursionStack.Contains(u))
            return true; // Cycle detected

        if (visited.Contains(u))
            return false;

        visited.Add(u);
        recursionStack.Add(u);

        if (_dependsOn.TryGetValue(u, out var neighbors))
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
