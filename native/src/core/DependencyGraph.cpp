#include "DependencyGraph.h"
#include <queue>

void DependencyGraph::addDependency(const CellAddress& dependent, const CellAddress& dependency) {
    uint64_t depKey = pack(dependent);
    uint64_t depOnKey = pack(dependency);

    m_dependencies[depKey].insert(depOnKey);
    m_dependents[depOnKey].insert(depKey);
}

void DependencyGraph::addRangeDependency(const CellAddress& dependent, const CellRange& range) {
    uint64_t depKey = pack(dependent);
    m_rangeDependencies[depKey].push_back(range);
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
            }
        }
        m_dependencies.erase(it);
    }

    // Remove range-level dependencies
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

// BFS to find all cells that need recalculation, in topological order
std::vector<CellAddress> DependencyGraph::getRecalcOrder(const CellAddress& changed) const {
    std::vector<CellAddress> order;
    std::unordered_set<uint64_t> visited;
    std::queue<uint64_t> queue;

    uint64_t startKey = pack(changed);
    auto it = m_dependents.find(startKey);
    if (it == m_dependents.end()) return order;

    for (uint64_t dep : it->second) {
        if (visited.find(dep) == visited.end()) {
            queue.push(dep);
            visited.insert(dep);
        }
    }

    while (!queue.empty()) {
        uint64_t current = queue.front();
        queue.pop();
        order.push_back(unpack(current));

        auto dit = m_dependents.find(current);
        if (dit != m_dependents.end()) {
            for (uint64_t dep : dit->second) {
                if (visited.find(dep) == visited.end()) {
                    queue.push(dep);
                    visited.insert(dep);
                }
            }
        }
    }

    return order;
}

bool DependencyGraph::hasCircularDependency(const CellAddress& cell) const {
    uint64_t cellKey = pack(cell);

    // First check: does this cell fall within any of its own range dependencies?
    // This catches self-referencing formulas like =SUM(A1:A100000) where the cell is in A1:A100000
    auto rangeIt = m_rangeDependencies.find(cellKey);
    if (rangeIt != m_rangeDependencies.end()) {
        for (const auto& range : rangeIt->second) {
            if (range.contains(cell)) {
                return true; // Cell references a range that contains itself
            }
        }
    }

    // Standard cycle detection via cell-level dependencies
    std::unordered_set<uint64_t> visited;
    return detectCycle(cellKey, cellKey, visited);
}

bool DependencyGraph::detectCycle(uint64_t start, uint64_t current,
                                   std::unordered_set<uint64_t>& visited) const {
    auto it = m_dependencies.find(current);
    if (it == m_dependencies.end()) return false;

    for (uint64_t dep : it->second) {
        if (dep == start) return true;
        if (visited.find(dep) == visited.end()) {
            visited.insert(dep);
            if (detectCycle(start, dep, visited)) return true;
        }
    }

    // Also check range dependencies of the current cell
    auto rangeIt = m_rangeDependencies.find(current);
    if (rangeIt != m_rangeDependencies.end()) {
        CellAddress startAddr = unpack(start);
        for (const auto& range : rangeIt->second) {
            if (range.contains(startAddr)) {
                return true; // Start cell falls within a range that current depends on
            }
        }
    }

    return false;
}

void DependencyGraph::shiftReferences(int atIndex, int delta, bool isRow) {
    // Helper to shift a packed address
    auto shiftKey = [&](uint64_t key) -> uint64_t {
        CellAddress addr = unpack(key);
        if (isRow) {
            if (addr.row >= atIndex) {
                addr.row += delta;
                if (addr.row < 0) return key; // don't shift into negative
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
}

void DependencyGraph::clear() {
    m_dependents.clear();
    m_dependencies.clear();
    m_rangeDependencies.clear();
}

void DependencyGraph::reserve(size_t cellCount) {
    m_dependents.reserve(cellCount);
    m_dependencies.reserve(cellCount);
}
