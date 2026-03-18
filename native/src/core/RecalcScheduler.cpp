#include "RecalcScheduler.h"
#include "Spreadsheet.h"
#include "FormulaEngine.h"
#include "DependencyGraph.h"
#include "ColumnStore.h"
#include <QtConcurrent>

RecalcScheduler::RecalcScheduler(Spreadsheet* spreadsheet)
    : m_spreadsheet(spreadsheet) {}

void RecalcScheduler::evaluateCell(const CellAddress& addr) {
    auto& store = m_spreadsheet->getColumnStore();
    if (store.getCellType(addr.row, addr.col) != CellDataType::Formula)
        return;

    // Thread-local FormulaEngine: no shared mutable state between threads
    thread_local FormulaEngine localEngine;
    localEngine.setSpreadsheet(m_spreadsheet);

    QString formula = store.getCellFormula(addr.row, addr.col);
    QVariant result = localEngine.evaluate(formula);
    store.setComputedValue(addr.row, addr.col, result);
}

void RecalcScheduler::evaluateLevel(const std::vector<CellAddress>& level) {
    if (level.empty()) return;

    if (static_cast<int>(level.size()) < m_parallelThreshold) {
        // Small level: single-threaded (avoid thread pool overhead)
        for (const auto& addr : level) {
            evaluateCell(addr);
        }
    } else {
        // Large level: parallel evaluation
        QtConcurrent::blockingMap(level, [this](const CellAddress& addr) {
            evaluateCell(addr);
        });
    }
}

std::vector<CellAddress> RecalcScheduler::recalculateDependents(const CellAddress& changed) {
    return recalculateDependents(std::vector<CellAddress>{changed});
}

std::vector<CellAddress> RecalcScheduler::recalculateDependents(
        const std::vector<CellAddress>& changed) {
    auto& depGraph = m_spreadsheet->getDependencyGraph();
    auto levels = depGraph.getRecalcLevels(changed);

    std::vector<CellAddress> allRecalculated;
    for (const auto& level : levels) {
        evaluateLevel(level);
        allRecalculated.insert(allRecalculated.end(), level.begin(), level.end());
    }
    return allRecalculated;
}

std::vector<CellAddress> RecalcScheduler::recalculateAll() {
    // Collect all formula cells as the "changed" set
    auto& store = m_spreadsheet->getColumnStore();
    std::vector<CellAddress> formulaCells;

    store.forEachCell([&](int row, int col, CellDataType type, const QVariant&) {
        if (type == CellDataType::Formula) {
            formulaCells.push_back(CellAddress(row, col));
        }
    });

    if (formulaCells.empty()) return {};

    // For full recalc, get topological levels from dependency graph
    auto& depGraph = m_spreadsheet->getDependencyGraph();
    auto levels = depGraph.getRecalcLevels(formulaCells);

    // If no levels returned (e.g., no deps registered), evaluate all in one pass
    if (levels.empty()) {
        evaluateLevel(formulaCells);
        return formulaCells;
    }

    std::vector<CellAddress> allRecalculated;
    for (const auto& level : levels) {
        evaluateLevel(level);
        allRecalculated.insert(allRecalculated.end(), level.begin(), level.end());
    }
    return allRecalculated;
}
