#ifndef FILTERENGINE_H
#define FILTERENGINE_H

#include <cstdint>
#include <vector>
#include <functional>
#include <QString>
#include <QVariant>
#include <QSet>
#include <bit>
#include "CellRange.h"

class ColumnStore;

// ============================================================================
// FilterEngine — Bitmap-based filtering for millions of rows
// ============================================================================
// Instead of calling setRowHidden() per row (20M API calls = 10+ seconds),
// we maintain a compact bitmap: 1 = passes filter, 0 = filtered out.
// 20M rows = 2.5 MB bitmap. Filter application: ~300ms parallel scan.
//
// The SpreadsheetModel/VirtualScrollBar use filteredToLogical() to map
// visible indices to data rows — zero QTableView row-hiding calls.
//
class FilterEngine {
public:
    FilterEngine();

    // Set the data source
    void setColumnStore(ColumnStore* store) { m_store = store; }

    // Set the range being filtered (e.g., entire data region)
    void setRange(int startRow, int endRow);

    // Apply a value filter on a column: only rows whose value is in `values` pass
    void applyValueFilter(int col, const QSet<QString>& values);

    // Apply a numeric condition filter (>, <, =, between, top N, etc.)
    enum class Condition { Eq, Neq, Lt, Gt, Lte, Gte, Between, TopN, BottomN };
    void applyConditionFilter(int col, Condition cond, double value1, double value2 = 0.0);

    // Apply a text filter (contains, starts with, ends with)
    enum class TextCondition { Contains, StartsWith, EndsWith, Equals };
    void applyTextFilter(int col, TextCondition cond, const QString& text, bool caseSensitive = false);

    // Apply a custom predicate filter
    void applyCustomFilter(int col, std::function<bool(const QVariant&)> predicate);

    // Clear filter on a specific column
    void clearFilter(int col);

    // Clear all filters
    void clearAllFilters();

    // Is any filter active?
    bool isFiltered() const { return m_isFiltered; }

    // Number of rows passing all filters
    int filteredRowCount() const { return static_cast<int>(m_filteredToLogical.size()); }

    // Map filtered row index → logical row
    int filteredToLogical(int filteredRow) const;

    // Map logical row → filtered row index (-1 if filtered out)
    int logicalToFiltered(int logicalRow) const;

    // Check if a specific logical row passes the filter
    bool rowPassesFilter(int logicalRow) const;

    // Get the raw bitmap (for external use, e.g., sort optimization)
    const std::vector<uint64_t>& getBitmap() const { return m_bitmap; }

private:
    ColumnStore* m_store = nullptr;
    int m_startRow = 0;
    int m_endRow = 0;
    bool m_isFiltered = false;

    // Bitmap: 1 bit per row in range. Bit set = passes filter.
    // Index: (logicalRow - m_startRow) / 64 → word, bit = (logicalRow - m_startRow) % 64
    std::vector<uint64_t> m_bitmap;

    // Precomputed mapping: filtered row index → logical row
    // Rebuilt after each filter change
    std::vector<int> m_filteredToLogical;

    // Per-column filter bitmaps (for combining multiple column filters)
    struct ColumnFilter {
        int col;
        std::vector<uint64_t> bitmap;
    };
    std::vector<ColumnFilter> m_columnFilters;

    void rebuildCombinedBitmap();
    void rebuildFilteredMapping();
    void initBitmapAllPass();

    // Helper: set bit for a row
    void setBit(std::vector<uint64_t>& bmp, int row) {
        int offset = row - m_startRow;
        bmp[offset / 64] |= (uint64_t(1) << (offset % 64));
    }

    // Helper: test bit for a row
    bool testBit(const std::vector<uint64_t>& bmp, int row) const {
        int offset = row - m_startRow;
        return (bmp[offset / 64] >> (offset % 64)) & 1;
    }
};

#endif // FILTERENGINE_H
