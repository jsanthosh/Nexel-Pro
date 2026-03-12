#include "DependencyGraph.h"
#include <queue>

void DependencyGraph::addDependency(const CellAddress& dependent, const CellAddress& dependency) {
    uint64_t depKey = pack(dependent);
    uint64_t depOnKey = pack(dependency);

    m_dependencies[depKey].insert(depOnKey);
    m_dependents[depOnKey].insert(depKey);
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
    std::unordered_set<uint64_t> visited;
    uint64_t cellKey = pack(cell);
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
    return false;
}

void DependencyGraph::clear() {
    m_dependents.clear();
    m_dependencies.clear();
}
