#ifndef DEPENDENCYGRAPH_H
#define DEPENDENCYGRAPH_H

#include <cstdint>
#include <unordered_map>
#include <unordered_set>
#include <vector>
#include "CellRange.h"

class DependencyGraph {
public:
    DependencyGraph() = default;

    void addDependency(const CellAddress& dependent, const CellAddress& dependency);
    // Add a range-level dependency (for large ranges where individual cells aren't tracked)
    void addRangeDependency(const CellAddress& dependent, const CellRange& range);
    void removeDependencies(const CellAddress& cell);
    std::vector<CellAddress> getDependents(const CellAddress& cell) const;
    std::vector<CellAddress> getDependencies(const CellAddress& cell) const;
    std::vector<CellAddress> getRecalcOrder(const CellAddress& changed) const;
    bool hasCircularDependency(const CellAddress& cell) const;
    // Shift packed addresses in-place for row/column insert/delete
    // isRow=true: shift row indices at/after atIndex by delta
    // isRow=false: shift column indices at/after atIndex by delta
    void shiftReferences(int atIndex, int delta, bool isRow);
    void clear();
    // Reserve capacity for known number of cells (reduces rehashing)
    void reserve(size_t cellCount);

private:
    static uint64_t pack(const CellAddress& addr) {
        return (static_cast<uint64_t>(static_cast<uint32_t>(addr.row)) << 32)
             | static_cast<uint64_t>(static_cast<uint32_t>(addr.col));
    }
    static CellAddress unpack(uint64_t key) {
        return CellAddress(static_cast<int>(key >> 32), static_cast<int>(key & 0xFFFFFFFF));
    }

    // cell -> set of cells that depend on it
    std::unordered_map<uint64_t, std::unordered_set<uint64_t>> m_dependents;
    // cell -> set of cells it depends on
    std::unordered_map<uint64_t, std::unordered_set<uint64_t>> m_dependencies;

    // Range-level dependencies: cell -> list of ranges it depends on (for large ranges)
    std::unordered_map<uint64_t, std::vector<CellRange>> m_rangeDependencies;

    bool detectCycle(uint64_t start, uint64_t current,
                     std::unordered_set<uint64_t>& visited) const;
};

#endif // DEPENDENCYGRAPH_H
