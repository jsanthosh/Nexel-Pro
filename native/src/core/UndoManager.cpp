#include "UndoManager.h"
#include "Spreadsheet.h"

static void restoreCell(Spreadsheet* sheet, const CellSnapshot& snap) {
    auto cell = sheet->getCell(snap.addr);
    if (snap.type == CellType::Formula) {
        cell->setFormula(snap.formula);
        QVariant result = sheet->getFormulaEngine().evaluate(snap.formula);
        cell->setComputedValue(result);
    } else if (snap.type == CellType::Empty) {
        cell->clear();
    } else {
        cell->setValue(snap.value);
    }
    cell->setStyle(snap.style);
}

// CellEditCommand
CellEditCommand::CellEditCommand(const CellSnapshot& before, const CellSnapshot& after)
    : m_before(before), m_after(after) {
}

void CellEditCommand::undo(Spreadsheet* sheet) {
    sheet->setAutoRecalculate(false);
    restoreCell(sheet, m_before);
    sheet->setAutoRecalculate(true);
}

void CellEditCommand::redo(Spreadsheet* sheet) {
    sheet->setAutoRecalculate(false);
    restoreCell(sheet, m_after);
    sheet->setAutoRecalculate(true);
}

// MultiCellEditCommand
MultiCellEditCommand::MultiCellEditCommand(const std::vector<CellSnapshot>& before,
                                           const std::vector<CellSnapshot>& after,
                                           const QString& desc)
    : m_before(before), m_after(after), m_description(desc) {
}

void MultiCellEditCommand::undo(Spreadsheet* sheet) {
    sheet->setAutoRecalculate(false);
    for (const auto& snap : m_before) {
        restoreCell(sheet, snap);
    }
    sheet->setAutoRecalculate(true);
}

void MultiCellEditCommand::redo(Spreadsheet* sheet) {
    sheet->setAutoRecalculate(false);
    for (const auto& snap : m_after) {
        restoreCell(sheet, snap);
    }
    sheet->setAutoRecalculate(true);
}

// StyleChangeCommand
StyleChangeCommand::StyleChangeCommand(const std::vector<CellSnapshot>& before,
                                       const std::vector<CellSnapshot>& after)
    : m_before(before), m_after(after) {
}

void StyleChangeCommand::undo(Spreadsheet* sheet) {
    for (const auto& snap : m_before) {
        auto cell = sheet->getCell(snap.addr);
        cell->setStyle(snap.style);
    }
}

void StyleChangeCommand::redo(Spreadsheet* sheet) {
    for (const auto& snap : m_after) {
        auto cell = sheet->getCell(snap.addr);
        cell->setStyle(snap.style);
    }
}

// InsertRowCommand
void InsertRowCommand::undo(Spreadsheet* sheet) {
    sheet->deleteRow(m_row, m_count);
}
void InsertRowCommand::redo(Spreadsheet* sheet) {
    sheet->insertRow(m_row, m_count);
}

// InsertColumnCommand
void InsertColumnCommand::undo(Spreadsheet* sheet) {
    sheet->deleteColumn(m_col, m_count);
}
void InsertColumnCommand::redo(Spreadsheet* sheet) {
    sheet->insertColumn(m_col, m_count);
    // Copy formatting from source column to newly inserted column(s)
    if (m_sourceCol >= 0) {
        // After insert, source col may have shifted right if inserted before it
        int srcCol = (m_col <= m_sourceCol) ? m_sourceCol + m_count : m_sourceCol;
        int maxRow = sheet->getRowCount();
        for (int r = 0; r < maxRow; ++r) {
            auto srcCell = sheet->getCellIfExists(r, srcCol);
            if (srcCell) {
                const CellStyle& srcStyle = srcCell->getStyle();
                for (int c = m_col; c < m_col + m_count; ++c) {
                    sheet->getCell(CellAddress(r, c))->setStyle(srcStyle);
                }
            }
        }
    }
}

// DeleteRowCommand
void DeleteRowCommand::undo(Spreadsheet* sheet) {
    sheet->setAutoRecalculate(false);
    sheet->insertRow(m_row, m_count);
    for (const auto& snap : m_deleted) {
        restoreCell(sheet, snap);
    }
    sheet->setAutoRecalculate(true);
}
void DeleteRowCommand::redo(Spreadsheet* sheet) {
    sheet->deleteRow(m_row, m_count);
}

// DeleteColumnCommand
void DeleteColumnCommand::undo(Spreadsheet* sheet) {
    sheet->setAutoRecalculate(false);
    sheet->insertColumn(m_col, m_count);
    for (const auto& snap : m_deleted) {
        restoreCell(sheet, snap);
    }
    sheet->setAutoRecalculate(true);
}
void DeleteColumnCommand::redo(Spreadsheet* sheet) {
    sheet->deleteColumn(m_col, m_count);
}

// UndoManager
void UndoManager::execute(std::unique_ptr<UndoCommand> cmd, Spreadsheet* sheet) {
    cmd->redo(sheet);
    m_undoStack.push_back(std::move(cmd));
    m_redoStack.clear();

    if (m_undoStack.size() > MAX_UNDO) {
        m_undoStack.erase(m_undoStack.begin());
    }
}

void UndoManager::pushCommand(std::unique_ptr<UndoCommand> cmd) {
    m_undoStack.push_back(std::move(cmd));
    m_redoStack.clear();

    if (m_undoStack.size() > MAX_UNDO) {
        m_undoStack.erase(m_undoStack.begin());
    }
}

void UndoManager::undo(Spreadsheet* sheet) {
    if (m_undoStack.empty()) return;
    auto cmd = std::move(m_undoStack.back());
    m_undoStack.pop_back();
    cmd->undo(sheet);
    m_redoStack.push_back(std::move(cmd));
    // Recalculate all formulas so dependents reflect the restored values
    sheet->recalculateAll();
}

void UndoManager::redo(Spreadsheet* sheet) {
    if (m_redoStack.empty()) return;
    auto cmd = std::move(m_redoStack.back());
    m_redoStack.pop_back();
    cmd->redo(sheet);
    m_undoStack.push_back(std::move(cmd));
    // Recalculate all formulas so dependents reflect the restored values
    sheet->recalculateAll();
}

bool UndoManager::canUndo() const {
    return !m_undoStack.empty();
}

bool UndoManager::canRedo() const {
    return !m_redoStack.empty();
}

QString UndoManager::undoText() const {
    if (m_undoStack.empty()) return QString();
    return m_undoStack.back()->description();
}

QString UndoManager::redoText() const {
    if (m_redoStack.empty()) return QString();
    return m_redoStack.back()->description();
}

CellAddress UndoManager::lastUndoTarget() const {
    if (m_redoStack.empty()) return CellAddress(0, 0);
    return m_redoStack.back()->targetCell();
}

CellAddress UndoManager::lastRedoTarget() const {
    if (m_undoStack.empty()) return CellAddress(0, 0);
    return m_undoStack.back()->targetCell();
}

void UndoManager::clear() {
    m_undoStack.clear();
    m_redoStack.clear();
}
