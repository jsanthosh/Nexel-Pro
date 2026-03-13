#include "UndoManager.h"
#include "Spreadsheet.h"

// --- CellSnapshot memory estimation ---
static size_t estimateSnapshotMemory(const CellSnapshot& snap) {
    size_t mem = sizeof(CellSnapshot);
    // QString internal storage: each stores size * 2 bytes (UTF-16) + overhead
    mem += static_cast<size_t>(snap.formula.size()) * 2;
    // QVariant: rough estimate based on type
    if (snap.value.typeId() == QMetaType::QString) {
        mem += static_cast<size_t>(snap.value.toString().size()) * 2;
    }
    // CellStyle strings
    mem += static_cast<size_t>(snap.style.fontName.size()) * 2;
    mem += static_cast<size_t>(snap.style.foregroundColor.size()) * 2;
    mem += static_cast<size_t>(snap.style.backgroundColor.size()) * 2;
    mem += static_cast<size_t>(snap.style.numberFormat.size()) * 2;
    mem += static_cast<size_t>(snap.style.currencyCode.size()) * 2;
    mem += static_cast<size_t>(snap.style.dateFormatId.size()) * 2;
    return mem;
}

static size_t estimateSnapshotVectorMemory(const std::vector<CellSnapshot>& snaps) {
    size_t mem = snaps.capacity() * sizeof(CellSnapshot);
    for (const auto& s : snaps) {
        mem += estimateSnapshotMemory(s);
    }
    return mem;
}

size_t CellSnapshot::estimateMemory() const {
    return estimateSnapshotMemory(*this);
}

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

size_t MultiCellEditCommand::estimateMemory() const {
    return sizeof(*this) + estimateSnapshotVectorMemory(m_before)
         + estimateSnapshotVectorMemory(m_after)
         + static_cast<size_t>(m_description.size()) * 2;
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

size_t StyleChangeCommand::estimateMemory() const {
    return sizeof(*this) + estimateSnapshotVectorMemory(m_before)
         + estimateSnapshotVectorMemory(m_after);
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
    // Restore heights of the re-inserted rows
    for (int i = 0; i < m_count && i < static_cast<int>(m_savedRowHeights.size()); ++i) {
        if (m_savedRowHeights[i] > 0)
            sheet->setRowHeight(m_row + i, m_savedRowHeights[i]);
    }
    sheet->setAutoRecalculate(true);
}
void DeleteRowCommand::redo(Spreadsheet* sheet) {
    // Save heights before deleting (for undo restore)
    m_savedRowHeights.clear();
    for (int i = 0; i < m_count; ++i)
        m_savedRowHeights.push_back(sheet->getRowHeight(m_row + i));
    sheet->deleteRow(m_row, m_count);
}

size_t DeleteRowCommand::estimateMemory() const {
    return sizeof(*this) + estimateSnapshotVectorMemory(m_deleted);
}

// DeleteColumnCommand
void DeleteColumnCommand::undo(Spreadsheet* sheet) {
    sheet->setAutoRecalculate(false);
    sheet->insertColumn(m_col, m_count);
    for (const auto& snap : m_deleted) {
        restoreCell(sheet, snap);
    }
    // Restore widths of the re-inserted columns
    for (int i = 0; i < m_count && i < static_cast<int>(m_savedColWidths.size()); ++i) {
        if (m_savedColWidths[i] > 0)
            sheet->setColumnWidth(m_col + i, m_savedColWidths[i]);
    }
    sheet->setAutoRecalculate(true);
}
void DeleteColumnCommand::redo(Spreadsheet* sheet) {
    // Save widths before deleting (for undo restore)
    m_savedColWidths.clear();
    for (int i = 0; i < m_count; ++i)
        m_savedColWidths.push_back(sheet->getColumnWidth(m_col + i));
    sheet->deleteColumn(m_col, m_count);
}

size_t DeleteColumnCommand::estimateMemory() const {
    return sizeof(*this) + estimateSnapshotVectorMemory(m_deleted);
}

// TableChangeCommand
void TableChangeCommand::undo(Spreadsheet* sheet) {
    sheet->getTablesRef() = m_before;
}

void TableChangeCommand::redo(Spreadsheet* sheet) {
    sheet->getTablesRef() = m_after;
}

// UndoManager
void UndoManager::enforceMemoryCap() {
    while (m_totalMemory > MAX_MEMORY_BYTES && !m_undoStack.empty()) {
        m_totalMemory -= m_undoStack.front()->estimateMemory();
        m_undoStack.pop_front();
    }
}

void UndoManager::execute(std::unique_ptr<UndoCommand> cmd, Spreadsheet* sheet) {
    cmd->redo(sheet);
    m_totalMemory += cmd->estimateMemory();
    m_undoStack.push_back(std::move(cmd));

    // Clear redo stack and account for memory
    for (const auto& c : m_redoStack) {
        m_totalMemory -= c->estimateMemory();
    }
    m_redoStack.clear();

    if (m_undoStack.size() > MAX_UNDO) {
        m_totalMemory -= m_undoStack.front()->estimateMemory();
        m_undoStack.pop_front();
    }
    enforceMemoryCap();
}

void UndoManager::pushCommand(std::unique_ptr<UndoCommand> cmd) {
    m_totalMemory += cmd->estimateMemory();
    m_undoStack.push_back(std::move(cmd));

    for (const auto& c : m_redoStack) {
        m_totalMemory -= c->estimateMemory();
    }
    m_redoStack.clear();

    if (m_undoStack.size() > MAX_UNDO) {
        m_totalMemory -= m_undoStack.front()->estimateMemory();
        m_undoStack.pop_front();
    }
    enforceMemoryCap();
}

void UndoManager::undo(Spreadsheet* sheet) {
    if (m_undoStack.empty()) return;
    auto cmd = std::move(m_undoStack.back());
    m_undoStack.pop_back();
    cmd->undo(sheet);
    // Memory moves from undo to redo stack — no net change
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

bool UndoManager::lastUndoIsStructural() const {
    if (m_redoStack.empty()) return false;
    return m_redoStack.back()->isStructural();
}

bool UndoManager::lastRedoIsStructural() const {
    if (m_undoStack.empty()) return false;
    return m_undoStack.back()->isStructural();
}

void UndoManager::clear() {
    m_undoStack.clear();
    m_redoStack.clear();
    m_totalMemory = 0;
}
