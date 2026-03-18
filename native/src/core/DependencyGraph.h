#ifndef DEPENDENCYGRAPH_H
#define DEPENDENCYGRAPH_H

#include <cstdint>
#include <unordered_map>
#include <unordered_set>
#include <vector>
#include <algorithm>
#include "CellRange.h"

// ============================================================================
// DependencyGraph — Topological recalc ordering with interval-tree ranges
// ============================================================================
// Key improvements over BFS-based predecessor:
//   1. Kahn's algorithm for correct topological evaluation order
//   2. Level-based output for parallel recalculation (cells in same level
//      are independent and can be evaluated concurrently)
//   3. Per-column interval tree for O(log N + K) range dependency lookup
//      instead of O(N) flat scan
//   4. Incremental shiftReferences — only touches keys >= atIndex
//
class DependencyGraph {
public:
    DependencyGraph() = default;

    void addDependency(const CellAddress& dependent, const CellAddress& dependency);
    void addRangeDependency(const CellAddress& dependent, const CellRange& range);
    void removeDependencies(const CellAddress& cell);

    std::vector<CellAddress> getDependents(const CellAddress& cell) const;
    std::vector<CellAddress> getDependencies(const CellAddress& cell) const;

    // Topological recalc order (Kahn's algorithm) — correct evaluation order
    std::vector<CellAddress> getRecalcOrder(const CellAddress& changed) const;

    // Level-based recalc order for parallel evaluation.
    // Level 0 = cells with no unresolved deps, can run in parallel.
    // Level 1 = cells depending only on level 0, etc.
    std::vector<std::vector<CellAddress>> getRecalcLevels(const CellAddress& changed) const;

    // Multi-cell version: dirty set → topological levels
    std::vector<std::vector<CellAddress>> getRecalcLevels(const std::vector<CellAddress>& changed) const;

    bool hasCircularDependency(const CellAddress& cell) const;

    // Incremental shift for row/column insert/delete
    void shiftReferences(int atIndex, int delta, bool isRow);

    void clear();
    void reserve(size_t cellCount);

    // Stats
    size_t edgeCount() const;

private:
    static uint64_t pack(const CellAddress& addr) {
        return (static_cast<uint64_t>(static_cast<uint32_t>(addr.row)) << 32)
             | static_cast<uint64_t>(static_cast<uint32_t>(addr.col));
    }
    static CellAddress unpack(uint64_t key) {
        return CellAddress(static_cast<int>(key >> 32), static_cast<int>(key & 0xFFFFFFFF));
    }

    // cell → set of cells that depend on it (forward edges)
    std::unordered_map<uint64_t, std::unordered_set<uint64_t>> m_dependents;
    // cell → set of cells it depends on (reverse edges)
    std::unordered_map<uint64_t, std::unordered_set<uint64_t>> m_dependencies;

    // Range-level dependencies: cell → list of ranges it depends on
    std::unordered_map<uint64_t, std::vector<CellRange>> m_rangeDependencies;

    // ---- Per-column interval index for range dependency lookups ----
    // For each column, sorted list of (startRow, endRow, dependentCellKey)
    // Enables O(log N + K) lookup: "which formulas depend on a range containing row R in column C?"
    struct RangeInterval {
        int startRow;
        int endRow;
        uint64_t dependentKey; // the formula cell that depends on this range
    };
    // column → sorted intervals (sorted by startRow for binary search)
    std::unordered_map<int, std::vector<RangeInterval>> m_columnIntervals;
    bool m_intervalsDirty = false;

    void rebuildIntervalIndex();
    void addToIntervalIndex(uint64_t depKey, const CellRange& range);
    void removeFromIntervalIndex(uint64_t depKey);

    // Find all formula cells that depend on a range containing (row, col)
    void queryRangeDependents(int row, int col, std::unordered_set<uint64_t>& results) const;

    // Collect all transitive dependents of a set of dirty cells
    void collectAllDependents(const std::unordered_set<uint64_t>& dirtyKeys,
                              std::unordered_set<uint64_t>& allAffected) const;

    bool detectCycle(uint64_t start, uint64_t current,
                     std::unordered_set<uint64_t>& visited) const;
};

#endif // DEPENDENCYGRAPH_H
