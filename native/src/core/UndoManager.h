#ifndef UNDOMANAGER_H
#define UNDOMANAGER_H

#include <QString>
#include <QVariant>
#include <vector>
#include <deque>
#include <memory>
#include "Cell.h"
#include "CellRange.h"
#include "TableStyle.h"

class Spreadsheet;

struct CellSnapshot {
    CellAddress addr;
    QVariant value;
    QString formula;
    CellStyle style;
    CellType type;

    // Estimate memory usage of this snapshot (for undo memory cap)
    size_t estimateMemory() const;
};

class UndoCommand {
public:
    virtual ~UndoCommand() = default;
    virtual void undo(Spreadsheet* sheet) = 0;
    virtual void redo(Spreadsheet* sheet) = 0;
    virtual QString description() const = 0;
    virtual CellAddress targetCell() const { return CellAddress(0, 0); }
    virtual bool isStructural() const { return false; }  // true for insert/delete row/col
    virtual size_t estimateMemory() const { return sizeof(*this); }
};

class CellEditCommand : public UndoCommand {
public:
    CellEditCommand(const CellSnapshot& before, const CellSnapshot& after);
    void undo(Spreadsheet* sheet) override;
    void redo(Spreadsheet* sheet) override;
    QString description() const override { return "Edit Cell"; }
    CellAddress targetCell() const override { return m_before.addr; }
    size_t estimateMemory() const override { return sizeof(*this) + m_before.estimateMemory() + m_after.estimateMemory(); }
private:
    CellSnapshot m_before;
    CellSnapshot m_after;
};

class MultiCellEditCommand : public UndoCommand {
public:
    MultiCellEditCommand(const std::vector<CellSnapshot>& before, const std::vector<CellSnapshot>& after, const QString& desc);
    void undo(Spreadsheet* sheet) override;
    void redo(Spreadsheet* sheet) override;
    QString description() const override { return m_description; }
    CellAddress targetCell() const override { return m_before.empty() ? CellAddress(0,0) : m_before.front().addr; }
    size_t estimateMemory() const override;
private:
    std::vector<CellSnapshot> m_before;
    std::vector<CellSnapshot> m_after;
    QString m_description;
};

class StyleChangeCommand : public UndoCommand {
public:
    StyleChangeCommand(const std::vector<CellSnapshot>& before, const std::vector<CellSnapshot>& after);
    void undo(Spreadsheet* sheet) override;
    void redo(Spreadsheet* sheet) override;
    QString description() const override { return "Change Style"; }
    size_t estimateMemory() const override;
private:
    std::vector<CellSnapshot> m_before;
    std::vector<CellSnapshot> m_after;
};

// Insert row(s) — undo deletes, redo inserts
class InsertRowCommand : public UndoCommand {
public:
    InsertRowCommand(int row, int count = 1) : m_row(row), m_count(count) {}
    void undo(Spreadsheet* sheet) override;
    void redo(Spreadsheet* sheet) override;
    QString description() const override { return "Insert Row"; }
    CellAddress targetCell() const override { return CellAddress(m_row, 0); }
    bool isStructural() const override { return true; }
private:
    int m_row;
    int m_count;
};

// Insert column(s) — undo deletes, redo inserts
class InsertColumnCommand : public UndoCommand {
public:
    InsertColumnCommand(int col, int count = 1, int sourceCol = -1)
        : m_col(col), m_count(count), m_sourceCol(sourceCol) {}
    void undo(Spreadsheet* sheet) override;
    void redo(Spreadsheet* sheet) override;
    QString description() const override { return "Insert Column"; }
    CellAddress targetCell() const override { return CellAddress(0, m_col); }
    bool isStructural() const override { return true; }
private:
    int m_col;
    int m_count;
    int m_sourceCol; // column to copy formatting from (-1 = none)
};

// Delete row(s) — saves deleted cells for undo restore
class DeleteRowCommand : public UndoCommand {
public:
    DeleteRowCommand(int row, int count, const std::vector<CellSnapshot>& deleted,
                     const std::vector<int>& savedRowHeights = {})
        : m_row(row), m_count(count), m_deleted(deleted), m_savedRowHeights(savedRowHeights) {}
    void undo(Spreadsheet* sheet) override;
    void redo(Spreadsheet* sheet) override;
    QString description() const override { return "Delete Row"; }
    CellAddress targetCell() const override { return CellAddress(m_row, 0); }
    bool isStructural() const override { return true; }
    size_t estimateMemory() const override;
private:
    int m_row;
    int m_count;
    std::vector<CellSnapshot> m_deleted;
    std::vector<int> m_savedRowHeights; // heights of deleted rows, for undo restore
};

// Delete column(s) — saves deleted cells for undo restore
class DeleteColumnCommand : public UndoCommand {
public:
    DeleteColumnCommand(int col, int count, const std::vector<CellSnapshot>& deleted,
                        const std::vector<int>& savedColWidths = {})
        : m_col(col), m_count(count), m_deleted(deleted), m_savedColWidths(savedColWidths) {}
    void undo(Spreadsheet* sheet) override;
    void redo(Spreadsheet* sheet) override;
    QString description() const override { return "Delete Column"; }
    CellAddress targetCell() const override { return CellAddress(0, m_col); }
    bool isStructural() const override { return true; }
    size_t estimateMemory() const override;
private:
    int m_col;
    int m_count;
    std::vector<CellSnapshot> m_deleted;
    std::vector<int> m_savedColWidths; // widths of deleted columns, for undo restore
};

// Table add/remove/style change — stores full table list snapshot
class TableChangeCommand : public UndoCommand {
public:
    TableChangeCommand(const std::vector<SpreadsheetTable>& before,
                       const std::vector<SpreadsheetTable>& after,
                       const CellRange& affectedRange,
                       const QString& desc)
        : m_before(before), m_after(after), m_affectedRange(affectedRange), m_description(desc) {}
    void undo(Spreadsheet* sheet) override;
    void redo(Spreadsheet* sheet) override;
    QString description() const override { return m_description; }
    CellAddress targetCell() const override { return m_affectedRange.getStart(); }
private:
    std::vector<SpreadsheetTable> m_before;
    std::vector<SpreadsheetTable> m_after;
    CellRange m_affectedRange;
    QString m_description;
};

// Compound command — groups multiple undo commands into a single undo step.
// Undo replays children in reverse order; redo replays in forward order.
class CompoundUndoCommand : public UndoCommand {
public:
    CompoundUndoCommand(const QString& desc) : m_description(desc) {}
    void addChild(std::unique_ptr<UndoCommand> cmd) { m_children.push_back(std::move(cmd)); }
    void undo(Spreadsheet* sheet) override {
        for (auto it = m_children.rbegin(); it != m_children.rend(); ++it)
            (*it)->undo(sheet);
    }
    void redo(Spreadsheet* sheet) override {
        for (auto& cmd : m_children)
            cmd->redo(sheet);
    }
    QString description() const override { return m_description; }
    CellAddress targetCell() const override {
        return m_children.empty() ? CellAddress(0, 0) : m_children.front()->targetCell();
    }
    bool isStructural() const override { return true; }
    size_t estimateMemory() const override {
        size_t total = sizeof(*this);
        for (const auto& cmd : m_children) total += cmd->estimateMemory();
        return total;
    }
private:
    std::vector<std::unique_ptr<UndoCommand>> m_children;
    QString m_description;
};

class UndoManager {
public:
    UndoManager() = default;

    void execute(std::unique_ptr<UndoCommand> cmd, Spreadsheet* sheet);
    void pushCommand(std::unique_ptr<UndoCommand> cmd);
    void undo(Spreadsheet* sheet);
    void redo(Spreadsheet* sheet);

    bool canUndo() const;
    bool canRedo() const;
    QString undoText() const;
    QString redoText() const;
    CellAddress lastUndoTarget() const;
    CellAddress lastRedoTarget() const;
    bool lastUndoIsStructural() const;
    bool lastRedoIsStructural() const;
    void clear();

    // Memory tracking
    size_t totalMemoryUsage() const { return m_totalMemory; }

private:
    void enforceMemoryCap();

    std::deque<std::unique_ptr<UndoCommand>> m_undoStack;
    std::deque<std::unique_ptr<UndoCommand>> m_redoStack;
    size_t m_totalMemory = 0;
    static constexpr size_t MAX_UNDO = 100;
    static constexpr size_t MAX_MEMORY_BYTES = 50 * 1024 * 1024; // 50MB cap
};

#endif // UNDOMANAGER_H
