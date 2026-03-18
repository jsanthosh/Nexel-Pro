#ifndef RECALCSCHEDULER_H
#define RECALCSCHEDULER_H

#include <vector>
#include <functional>
#include "CellRange.h"

class Spreadsheet;
class DependencyGraph;

// ============================================================================
// RecalcScheduler — Level-parallel multi-threaded recalculation
// ============================================================================
// Uses DependencyGraph::getRecalcLevels() to get topological levels,
// then evaluates each level in parallel using QtConcurrent.
//
// Cells within the same level have no mutual dependencies and can be
// evaluated concurrently on separate threads.
//
// Thread safety: Each thread uses a thread_local FormulaEngine instance
// and writes only to its own cell's computed value slot in the ColumnStore
// (different cells = different memory locations = no data races).
//
class RecalcScheduler {
public:
    explicit RecalcScheduler(Spreadsheet* spreadsheet);

    // Recalculate all dependents of a changed cell using parallel levels
    // Returns list of all cells that were recalculated
    std::vector<CellAddress> recalculateDependents(const CellAddress& changed);

    // Recalculate dependents of multiple changed cells
    std::vector<CellAddress> recalculateDependents(const std::vector<CellAddress>& changed);

    // Full recalculation of all formula cells using parallel levels
    std::vector<CellAddress> recalculateAll();

    // Minimum cells per level to engage parallel evaluation
    // Below this, single-threaded is faster (avoids thread overhead)
    void setParallelThreshold(int threshold) { m_parallelThreshold = threshold; }

private:
    Spreadsheet* m_spreadsheet;
    int m_parallelThreshold = 64;

    void evaluateCell(const CellAddress& addr);
    void evaluateLevel(const std::vector<CellAddress>& level);
};

#endif // RECALCSCHEDULER_H
