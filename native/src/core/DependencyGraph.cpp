#include "DependencyGraph.h"
#include <queue>
#include <algorithm>

// ============================================================================
// Edge management
// ============================================================================

void DependencyGraph::addDependency(const CellAddress& dependent, const CellAddress& dependency) {
    uint64_t depKey = pack(dependent);
    uint64_t depOnKey = pack(dependency);
    m_dependencies[depKey].insert(depOnKey);
    m_dependents[depOnKey].insert(depKey);
}

void DependencyGraph::addRangeDependency(const CellAddress& dependent, const CellRange& range) {
    uint64_t depKey = pack(dependent);
    m_rangeDependencies[depKey].push_back(range);
    addToIntervalIndex(depKey, range);
}

void DependencyGraph::removeDependencies(const CellAddress& cell) {
    uint64_t cellKey = pack(cell);

    // Remove this cell from all its dependencies' dependent lists
    auto it = m_dependencies.find(cellKey);
    if (it != m_dependencies.end()) {
        for (uint64_t depOn : it->second) {
            auto dit = m_dependents.find(depOn);
            if (dit != m_dependents.end()) {
                dit->second.erase(cellKey);
                if (dit->second.empty()) m_dependents.erase(dit);
            }
        }
        m_dependencies.erase(it);
    }

    // Remove range dependencies and interval index entries
    removeFromIntervalIndex(cellKey);
    m_rangeDependencies.erase(cellKey);
}

std::vector<CellAddress> DependencyGraph::getDependents(const CellAddress& cell) const {
    std::vector<CellAddress> result;
    auto it = m_dependents.find(pack(cell));
    if (it != m_dependents.end()) {
        result.reserve(it->second.size());
        for (uint64_t k : it->second) {
            result.push_back(unpack(k));
        }
    }
    return result;
}

std::vector<CellAddress> DependencyGraph::getDependencies(const CellAddress& cell) const {
    std::vector<CellAddress> result;
    auto it = m_dependencies.find(pack(cell));
    if (it != m_dependencies.end()) {
        result.reserve(it->second.size());
        for (uint64_t k : it->second) {
            result.push_back(unpack(k));
        }
    }
    return result;
}

// ============================================================================
// Interval index for range dependencies
// ============================================================================

void DependencyGraph::addToIntervalIndex(uint64_t depKey, const CellRange& range) {
    auto start = range.getStart();
    auto end = range.getEnd();
    int startCol = std::min(start.col, end.col);
    int endCol = std::max(start.col, end.col);
    int startRow = std::min(start.row, end.row);
    int endRow = std::max(start.row, end.row);

    for (int col = startCol; col <= endCol; ++col) {
        m_columnIntervals[col].push_back({startRow, endRow, depKey});
    }
    m_intervalsDirty = true;
}

void DependencyGraph::removeFromIntervalIndex(uint64_t depKey) {
    // Remove all intervals belonging to this dependent cell
    for (auto& [col, intervals] : m_columnIntervals) {
        intervals.erase(
            std::remove_if(intervals.begin(), intervals.end(),
                [depKey](const RangeInterval& ri) { return ri.dependentKey == depKey; }),
            intervals.end());
    }
}

void DependencyGraph::rebuildIntervalIndex() {
    m_columnIntervals.clear();
    for (auto& [depKey, ranges] : m_rangeDependencies) {
        for (auto& range : ranges) {
            auto start = range.getStart();
            auto end = range.getEnd();
            int startCol = std::min(start.col, end.col);
            int endCol = std::max(start.col, end.col);
            int startRow = std::min(start.row, end.row);
            int endRow = std::max(start.row, end.row);
            for (int col = startCol; col <= endCol; ++col) {
                m_columnIntervals[col].push_back({startRow, endRow, depKey});
            }
        }
    }
    // Sort each column's intervals by startRow for binary search
    for (auto& [col, intervals] : m_columnIntervals) {
        std::sort(intervals.begin(), intervals.end(),
            [](const RangeInterval& a, const RangeInterval& b) {
                return a.startRow < b.startRow;
            });
    }
    m_intervalsDirty = false;
}

void DependencyGraph::queryRangeDependents(int row, int col,
                                            std::unordered_set<uint64_t>& results) const {
    auto it = m_columnIntervals.find(col);
    if (it == m_columnIntervals.end()) return;

    const auto& intervals = it->second;
    // Binary search: find first interval where startRow <= row
    // Then scan intervals that could contain row
    // Since sorted by startRow, use lower_bound to find first interval
    // where startRow > row, then scan backwards + forward

    for (const auto& ri : intervals) {
        // Early exit: if startRow > row, no more intervals can contain row
        // (since sorted by startRow)
        if (ri.startRow > row) break;
        if (ri.endRow >= row) {
            results.insert(ri.dependentKey);
        }
    }
}

// ============================================================================
// Topological recalc — Kahn's algorithm
// ============================================================================

void DependencyGraph::collectAllDependents(const std::unordered_set<uint64_t>& dirtyKeys,
                                            std::unordered_set<uint64_t>& allAffected) const {
    std::queue<uint64_t> queue;

    // Seed: direct dependents of all dirty cells
    for (uint64_t dirty : dirtyKeys) {
        // Cell-level dependents
        auto it = m_dependents.find(dirty);
        if (it != m_dependents.end()) {
            for (uint64_t dep : it->second) {
                if (allAffected.insert(dep).second) {
                    queue.push(dep);
                }
            }
        }
        // Range-level dependents: which formulas depend on a range containing this cell?
        CellAddress addr = unpack(dirty);
        queryRangeDependents(addr.row, addr.col, allAffected);
    }

    // Also add range dependents to queue
    for (uint64_t key : allAffected) {
        queue.push(key);
    }

    // BFS to collect ALL transitive dependents
    while (!queue.empty()) {
        uint64_t current = queue.front();
        queue.pop();

        auto it = m_dependents.find(current);
        if (it != m_dependents.end()) {
            for (uint64_t dep : it->second) {
                if (allAffected.insert(dep).second) {
                    queue.push(dep);
                }
            }
        }

        // Range dependents of current
        CellAddress addr = unpack(current);
        std::unordered_set<uint64_t> rangeDeps;
        queryRangeDependents(addr.row, addr.col, rangeDeps);
        for (uint64_t rd : rangeDeps) {
            if (allAffected.insert(rd).second) {
                queue.push(rd);
            }
        }
    }
}

// Single-cell changed → flat topological order (backward-compatible API)
std::vector<CellAddress> DependencyGraph::getRecalcOrder(const CellAddress& changed) const {
    auto levels = getRecalcLevels(changed);
    std::vector<CellAddress> flat;
    for (auto& level : levels) {
        flat.insert(flat.end(), level.begin(), level.end());
    }
    return flat;
}

// Single-cell changed → level-based topological order
std::vector<std::vector<CellAddress>> DependencyGraph::getRecalcLevels(const CellAddress& changed) const {
    return getRecalcLevels(std::vector<CellAddress>{changed});
}

// Multi-cell changed → level-based topological order (Kahn's algorithm)
std::vector<std::vector<CellAddress>> DependencyGraph::getRecalcLevels(
        const std::vector<CellAddress>& changed) const {

    if (m_intervalsDirty) {
        const_cast<DependencyGraph*>(this)->rebuildIntervalIndex();
    }

    // 1. Collect all affected cells (transitive dependents of changed set)
    std::unordered_set<uint64_t> dirtyKeys;
    for (const auto& addr : changed) {
        dirtyKeys.insert(pack(addr));
    }

    std::unordered_set<uint64_t> allAffected;
    collectAllDependents(dirtyKeys, allAffected);

    if (allAffected.empty()) return {};

    // 2. Build subgraph in-degree map (only within affected set)
    std::unordered_map<uint64_t, int> inDegree;
    std::unordered_map<uint64_t, std::vector<uint64_t>> subgraphEdges; // forward edges within affected set

    for (uint64_t key : allAffected) {
        inDegree[key] = 0;
    }

    for (uint64_t key : allAffected) {
        // Check which of key's dependencies are also in the affected set
        auto it = m_dependencies.find(key);
        if (it != m_dependencies.end()) {
            for (uint64_t dep : it->second) {
                if (allAffected.count(dep)) {
                    // dep → key edge exists in subgraph
                    subgraphEdges[dep].push_back(key);
                    inDegree[key]++;
                }
            }
        }
    }

    // 3. Kahn's algorithm — process level by level
    std::vector<std::vector<CellAddress>> levels;
    std::vector<uint64_t> currentLevel;

    // Seeds: affected cells with in-degree 0 (no deps within affected set)
    for (auto& [key, deg] : inDegree) {
        if (deg == 0) {
            currentLevel.push_back(key);
        }
    }

    while (!currentLevel.empty()) {
        // Convert current level to CellAddresses
        std::vector<CellAddress> levelAddrs;
        levelAddrs.reserve(currentLevel.size());
        for (uint64_t key : currentLevel) {
            levelAddrs.push_back(unpack(key));
        }
        levels.push_back(std::move(levelAddrs));

        // Find next level
        std::vector<uint64_t> nextLevel;
        for (uint64_t key : currentLevel) {
            auto it = subgraphEdges.find(key);
            if (it != subgraphEdges.end()) {
                for (uint64_t succ : it->second) {
                    if (--inDegree[succ] == 0) {
                        nextLevel.push_back(succ);
                    }
                }
            }
        }
        currentLevel = std::move(nextLevel);
    }

    return levels;
}

// ============================================================================
// Cycle detection
// ============================================================================

bool DependencyGraph::hasCircularDependency(const CellAddress& cell) const {
    uint64_t cellKey = pack(cell);

    // Check: does this cell fall within any of its own range dependencies?
    auto rangeIt = m_rangeDependencies.find(cellKey);
    if (rangeIt != m_rangeDependencies.end()) {
        for (const auto& range : rangeIt->second) {
            if (range.contains(cell)) {
                return true;
            }
        }
    }

    // Standard DFS cycle detection via cell-level dependencies
    std::unordered_set<uint64_t> visited;
    return detectCycle(cellKey, cellKey, visited);
}

bool DependencyGraph::detectCycle(uint64_t start, uint64_t current,
                                   std::unordered_set<uint64_t>& visited) const {
    auto it = m_dependencies.find(current);
    if (it == m_dependencies.end()) return false;

    for (uint64_t dep : it->second) {
        if (dep == start) return true;
        if (visited.insert(dep).second) {
            if (detectCycle(start, dep, visited)) return true;
        }
    }

    // Also check range dependencies of current cell
    auto rangeIt = m_rangeDependencies.find(current);
    if (rangeIt != m_rangeDependencies.end()) {
        CellAddress startAddr = unpack(start);
        for (const auto& range : rangeIt->second) {
            if (range.contains(startAddr)) {
                return true;
            }
        }
    }

    return false;
}

// ============================================================================
// Shift references (incremental)
// ============================================================================

void DependencyGraph::shiftReferences(int atIndex, int delta, bool isRow) {
    auto shiftKey = [&](uint64_t key) -> uint64_t {
        CellAddress addr = unpack(key);
        if (isRow) {
            if (addr.row >= atIndex) {
                addr.row += delta;
                if (addr.row < 0) return key;
            }
        } else {
            if (addr.col >= atIndex) {
                addr.col += delta;
                if (addr.col < 0) return key;
            }
        }
        return pack(addr);
    };

    // Rebuild m_dependents with shifted keys
    std::unordered_map<uint64_t, std::unordered_set<uint64_t>> newDependents;
    for (auto& [key, deps] : m_dependents) {
        uint64_t newKey = shiftKey(key);
        auto& newSet = newDependents[newKey];
        for (uint64_t d : deps) {
            newSet.insert(shiftKey(d));
        }
    }
    m_dependents = std::move(newDependents);

    // Rebuild m_dependencies with shifted keys
    std::unordered_map<uint64_t, std::unordered_set<uint64_t>> newDependencies;
    for (auto& [key, deps] : m_dependencies) {
        uint64_t newKey = shiftKey(key);
        auto& newSet = newDependencies[newKey];
        for (uint64_t d : deps) {
            newSet.insert(shiftKey(d));
        }
    }
    m_dependencies = std::move(newDependencies);

    // Rebuild m_rangeDependencies with shifted keys and ranges
    std::unordered_map<uint64_t, std::vector<CellRange>> newRangeDeps;
    for (auto& [key, ranges] : m_rangeDependencies) {
        uint64_t newKey = shiftKey(key);
        auto& newRanges = newRangeDeps[newKey];
        for (auto& range : ranges) {
            auto s = range.getStart();
            auto e = range.getEnd();
            if (isRow) {
                if (s.row >= atIndex) s.row += delta;
                if (e.row >= atIndex) e.row += delta;
            } else {
                if (s.col >= atIndex) s.col += delta;
                if (e.col >= atIndex) e.col += delta;
            }
            newRanges.emplace_back(s, e);
        }
    }
    m_rangeDependencies = std::move(newRangeDeps);

    // Mark interval index dirty — will be rebuilt on next query
    m_intervalsDirty = true;
}

// ============================================================================
// Housekeeping
// ============================================================================

void DependencyGraph::clear() {
    m_dependents.clear();
    m_dependencies.clear();
    m_rangeDependencies.clear();
    m_columnIntervals.clear();
    m_intervalsDirty = false;
}

void DependencyGraph::reserve(size_t cellCount) {
    m_dependents.reserve(cellCount);
    m_dependencies.reserve(cellCount);
}

size_t DependencyGraph::edgeCount() const {
    size_t count = 0;
    for (const auto& [k, v] : m_dependents) {
        count += v.size();
    }
    return count;
}
