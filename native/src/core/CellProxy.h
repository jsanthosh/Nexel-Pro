#ifndef CELLPROXY_H
#define CELLPROXY_H

#include <QString>
#include <QVariant>
#include "Cell.h"
#include "CellRange.h"
#include "ColumnStore.h"
#include "StyleTable.h"
#include "StringPool.h"

// ============================================================================
// CellProxy — Lightweight accessor for ColumnStore cells
// ============================================================================
// Replaces shared_ptr<Cell> as the public API for cell access.
// Returned BY VALUE (16 bytes: pointer + row + col) — no heap allocation.
// Provides the same interface as Cell so existing code can migrate incrementally.
//
// Usage:
//   CellProxy cell = spreadsheet.cell(row, col);
//   QVariant value = cell.getValue();
//   cell.setValue(42);
//
class CellProxy {
public:
    CellProxy() : m_store(nullptr), m_row(-1), m_col(-1) {}
    CellProxy(ColumnStore* store, int row, int col)
        : m_store(store), m_row(row), m_col(col) {}

    bool isValid() const { return m_store != nullptr && m_row >= 0 && m_col >= 0; }
    bool isEmpty() const { return !m_store || !m_store->hasCell(m_row, m_col); }

    int row() const { return m_row; }
    int col() const { return m_col; }

    // ---- Value access (matches Cell interface) ----

    QVariant getValue() const {
        if (!m_store) return QVariant();
        // For formula cells, always return computed value (never the raw formula string)
        auto type = m_store->getCellType(m_row, m_col);
        if (type == CellDataType::Formula) {
            return m_store->getComputedValue(m_row, m_col);
        }
        return m_store->getCellValue(m_row, m_col);
    }

    void setValue(const QVariant& value) {
        if (m_store) m_store->setCellValue(m_row, m_col, value);
    }

    // ---- Formula ----

    QString getFormula() const {
        if (!m_store) return QString();
        return m_store->getCellFormula(m_row, m_col);
    }

    void setFormula(const QString& formula) {
        if (m_store) m_store->setCellFormula(m_row, m_col, formula);
    }

    bool hasFormula() const {
        return m_store && m_store->getCellType(m_row, m_col) == CellDataType::Formula;
    }

    // ---- Type ----

    CellType getType() const {
        if (!m_store) return CellType::Empty;
        switch (m_store->getCellType(m_row, m_col)) {
            case CellDataType::Empty:   return CellType::Empty;
            case CellDataType::Double:  return CellType::Number;
            case CellDataType::String:  return CellType::Text;
            case CellDataType::Formula: return CellType::Formula;
            case CellDataType::Boolean: return CellType::Boolean;
            case CellDataType::Error:   return CellType::Error;
            case CellDataType::Date:    return CellType::Date;
        }
        return CellType::Empty;
    }

    // ---- Style ----

    const CellStyle& getStyle() const {
        if (!m_store) return Cell::defaultStyle();
        uint16_t idx = m_store->getCellStyleIndex(m_row, m_col);
        return StyleTable::instance().get(idx);
    }

    void setStyle(const CellStyle& style) {
        if (!m_store) return;
        uint16_t idx = StyleTable::instance().intern(style);
        m_store->setCellStyle(m_row, m_col, idx);
    }

    bool hasCustomStyle() const {
        if (!m_store) return false;
        return m_store->getCellStyleIndex(m_row, m_col) != 0;
    }

    uint16_t getStyleIndex() const {
        if (!m_store) return 0;
        return m_store->getCellStyleIndex(m_row, m_col);
    }

    // ---- Computed value ----

    QVariant getComputedValue() const {
        if (!m_store) return QVariant();
        return m_store->getComputedValue(m_row, m_col);
    }

    void setComputedValue(const QVariant& value) {
        if (m_store) m_store->setComputedValue(m_row, m_col, value);
    }

    // ---- Dirty flag ----
    // For now, delegated to a simple check on the store
    bool isDirty() const { return false; } // TODO: integrate with ColumnStore flags bitmap
    void setDirty(bool) {} // TODO

    // ---- Error ----

    bool hasError() const {
        return m_store && m_store->getCellType(m_row, m_col) == CellDataType::Error;
    }

    QString getError() const {
        if (!m_store) return QString();
        if (m_store->getCellType(m_row, m_col) != CellDataType::Error) return QString();
        return m_store->getCellValue(m_row, m_col).toString();
    }

    void setError(const QString& error) {
        if (!m_store) return;
        uint32_t id = StringPool::instance().intern(error);
        auto* column = m_store->getOrCreateColumn(m_col);
        auto* chunk = column->getOrCreateChunk(m_row);
        chunk->setError(m_row - chunk->baseRow, id);
    }

    // ---- Comments ----

    QString getComment() const {
        if (!m_store) return QString();
        return m_store->getCellComment(m_row, m_col);
    }

    void setComment(const QString& comment) {
        if (m_store) m_store->setCellComment(m_row, m_col, comment);
    }

    bool hasComment() const {
        if (!m_store) return false;
        return m_store->hasCellComment(m_row, m_col);
    }

    // ---- Hyperlinks ----

    QString getHyperlink() const {
        if (!m_store) return QString();
        return m_store->getCellHyperlink(m_row, m_col);
    }

    void setHyperlink(const QString& url) {
        if (m_store) m_store->setCellHyperlink(m_row, m_col, url);
    }

    bool hasHyperlink() const {
        if (!m_store) return false;
        return m_store->hasCellHyperlink(m_row, m_col);
    }

    // ---- Spill tracking ----

    CellAddress getSpillParent() const {
        if (!m_store) return CellAddress(-1, -1);
        return m_store->getSpillParent(m_row, m_col);
    }

    void setSpillParent(const CellAddress& parent) {
        if (m_store) m_store->setSpillParent(m_row, m_col, parent);
    }

    bool isSpillCell() const {
        if (!m_store) return false;
        return m_store->isSpillCell(m_row, m_col);
    }

    void clearSpillParent() {
        if (m_store) m_store->clearSpillParent(m_row, m_col);
    }

    // ---- Fast bulk setters ----

    void setValueDirect(double num) {
        if (m_store) m_store->setCellNumeric(m_row, m_col, num);
    }

    void setValueDirect(const QString& text) {
        if (m_store) m_store->setCellString(m_row, m_col, text);
    }

    // ---- Utilities ----

    void clear() {
        if (m_store) m_store->removeCell(m_row, m_col);
    }

    QString toString() const {
        QVariant val = getValue();
        return val.isValid() ? val.toString() : QString();
    }

    // ---- Arrow operator (transparent migration from shared_ptr<Cell>) ----
    // Allows cell->getValue() to work the same as cell.getValue()
    CellProxy* operator->() { return this; }
    const CellProxy* operator->() const { return this; }

    // ---- Comparison (for compatibility with code that checks nullptr) ----
    explicit operator bool() const { return isValid() && !isEmpty(); }

private:
    ColumnStore* m_store;
    int m_row;
    int m_col;
};

#endif // CELLPROXY_H
