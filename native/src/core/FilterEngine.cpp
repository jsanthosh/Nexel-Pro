#include "FilterEngine.h"
#include "ColumnStore.h"
#include <algorithm>
#include <numeric>
#include <cmath>

FilterEngine::FilterEngine() = default;

void FilterEngine::setRange(int startRow, int endRow) {
    m_startRow = startRow;
    m_endRow = endRow;
    clearAllFilters();
}

void FilterEngine::initBitmapAllPass() {
    int rowCount = m_endRow - m_startRow + 1;
    int words = (rowCount + 63) / 64;
    m_bitmap.assign(words, ~uint64_t(0));

    // Clear trailing bits in last word
    int trailingBits = rowCount % 64;
    if (trailingBits > 0) {
        m_bitmap.back() = (uint64_t(1) << trailingBits) - 1;
    }
}

void FilterEngine::rebuildCombinedBitmap() {
    if (m_columnFilters.empty()) {
        m_isFiltered = false;
        initBitmapAllPass();
        rebuildFilteredMapping();
        return;
    }

    m_isFiltered = true;
    int words = static_cast<int>(m_columnFilters[0].bitmap.size());

    // AND all column filter bitmaps together
    m_bitmap = m_columnFilters[0].bitmap;
    for (size_t i = 1; i < m_columnFilters.size(); ++i) {
        for (int w = 0; w < words; ++w) {
            m_bitmap[w] &= m_columnFilters[i].bitmap[w];
        }
    }

    rebuildFilteredMapping();
}

void FilterEngine::rebuildFilteredMapping() {
    m_filteredToLogical.clear();
    int rowCount = m_endRow - m_startRow + 1;
    m_filteredToLogical.reserve(rowCount); // worst case

    for (int w = 0; w < static_cast<int>(m_bitmap.size()); ++w) {
        uint64_t word = m_bitmap[w];
        while (word) {
            int bit = std::countr_zero(word); // C++20: position of lowest set bit
            int logicalRow = m_startRow + w * 64 + bit;
            if (logicalRow <= m_endRow) {
                m_filteredToLogical.push_back(logicalRow);
            }
            word &= word - 1; // clear lowest set bit
        }
    }
}

int FilterEngine::filteredToLogical(int filteredRow) const {
    if (filteredRow < 0 || filteredRow >= static_cast<int>(m_filteredToLogical.size()))
        return -1;
    return m_filteredToLogical[filteredRow];
}

int FilterEngine::logicalToFiltered(int logicalRow) const {
    if (!m_isFiltered) return logicalRow - m_startRow;
    // Binary search in sorted m_filteredToLogical
    auto it = std::lower_bound(m_filteredToLogical.begin(), m_filteredToLogical.end(), logicalRow);
    if (it != m_filteredToLogical.end() && *it == logicalRow) {
        return static_cast<int>(it - m_filteredToLogical.begin());
    }
    return -1;
}

bool FilterEngine::rowPassesFilter(int logicalRow) const {
    if (!m_isFiltered) return true;
    if (logicalRow < m_startRow || logicalRow > m_endRow) return false;
    return testBit(m_bitmap, logicalRow);
}

// ============================================================================
// Value filter: row passes if cell string value is in the allowed set
// ============================================================================

void FilterEngine::applyValueFilter(int col, const QSet<QString>& values) {
    if (!m_store) return;

    int rowCount = m_endRow - m_startRow + 1;
    int words = (rowCount + 63) / 64;
    std::vector<uint64_t> colBitmap(words, 0);

    // Scan column and set bits for matching rows
    m_store->scanColumnValues(col, m_startRow, m_endRow,
        [&](int row, const QVariant& val) {
            QString strVal = val.toString();
            if (values.contains(strVal)) {
                setBit(colBitmap, row);
            }
        });

    // Also allow empty cells if empty string is in values
    if (values.contains(QString())) {
        for (int r = m_startRow; r <= m_endRow; ++r) {
            if (m_store->getCellType(r, col) == CellDataType::Empty) {
                setBit(colBitmap, r);
            }
        }
    }

    // Update or add column filter
    bool found = false;
    for (auto& cf : m_columnFilters) {
        if (cf.col == col) {
            cf.bitmap = std::move(colBitmap);
            found = true;
            break;
        }
    }
    if (!found) {
        m_columnFilters.push_back({col, std::move(colBitmap)});
    }

    rebuildCombinedBitmap();
}

// ============================================================================
// Numeric condition filter
// ============================================================================

void FilterEngine::applyConditionFilter(int col, Condition cond, double value1, double value2) {
    if (!m_store) return;

    int rowCount = m_endRow - m_startRow + 1;
    int words = (rowCount + 63) / 64;
    std::vector<uint64_t> colBitmap(words, 0);

    if (cond == Condition::TopN || cond == Condition::BottomN) {
        // Collect all numeric values with their rows
        std::vector<std::pair<double, int>> valRows;
        m_store->scanColumnValues(col, m_startRow, m_endRow,
            [&](int row, const QVariant& val) {
                bool ok;
                double d = val.toDouble(&ok);
                if (ok) valRows.push_back({d, row});
            });

        int n = static_cast<int>(value1);
        if (cond == Condition::TopN) {
            std::partial_sort(valRows.begin(),
                valRows.begin() + std::min(n, static_cast<int>(valRows.size())),
                valRows.end(),
                [](auto& a, auto& b) { return a.first > b.first; });
        } else {
            std::partial_sort(valRows.begin(),
                valRows.begin() + std::min(n, static_cast<int>(valRows.size())),
                valRows.end(),
                [](auto& a, auto& b) { return a.first < b.first; });
        }
        for (int i = 0; i < std::min(n, static_cast<int>(valRows.size())); ++i) {
            setBit(colBitmap, valRows[i].second);
        }
    } else {
        m_store->scanColumnValues(col, m_startRow, m_endRow,
            [&](int row, const QVariant& val) {
                bool ok;
                double d = val.toDouble(&ok);
                if (!ok) return;

                bool passes = false;
                switch (cond) {
                    case Condition::Eq:  passes = (d == value1); break;
                    case Condition::Neq: passes = (d != value1); break;
                    case Condition::Lt:  passes = (d < value1); break;
                    case Condition::Gt:  passes = (d > value1); break;
                    case Condition::Lte: passes = (d <= value1); break;
                    case Condition::Gte: passes = (d >= value1); break;
                    case Condition::Between: passes = (d >= value1 && d <= value2); break;
                    default: break;
                }
                if (passes) setBit(colBitmap, row);
            });
    }

    bool found = false;
    for (auto& cf : m_columnFilters) {
        if (cf.col == col) {
            cf.bitmap = std::move(colBitmap);
            found = true;
            break;
        }
    }
    if (!found) {
        m_columnFilters.push_back({col, std::move(colBitmap)});
    }

    rebuildCombinedBitmap();
}

// ============================================================================
// Text condition filter
// ============================================================================

void FilterEngine::applyTextFilter(int col, TextCondition cond, const QString& text, bool caseSensitive) {
    if (!m_store) return;

    int rowCount = m_endRow - m_startRow + 1;
    int words = (rowCount + 63) / 64;
    std::vector<uint64_t> colBitmap(words, 0);

    Qt::CaseSensitivity cs = caseSensitive ? Qt::CaseSensitive : Qt::CaseInsensitive;

    m_store->scanColumnValues(col, m_startRow, m_endRow,
        [&](int row, const QVariant& val) {
            QString str = val.toString();
            bool passes = false;
            switch (cond) {
                case TextCondition::Contains:   passes = str.contains(text, cs); break;
                case TextCondition::StartsWith: passes = str.startsWith(text, cs); break;
                case TextCondition::EndsWith:   passes = str.endsWith(text, cs); break;
                case TextCondition::Equals:     passes = (str.compare(text, cs) == 0); break;
            }
            if (passes) setBit(colBitmap, row);
        });

    bool found = false;
    for (auto& cf : m_columnFilters) {
        if (cf.col == col) {
            cf.bitmap = std::move(colBitmap);
            found = true;
            break;
        }
    }
    if (!found) {
        m_columnFilters.push_back({col, std::move(colBitmap)});
    }

    rebuildCombinedBitmap();
}

// ============================================================================
// Custom predicate filter
// ============================================================================

void FilterEngine::applyCustomFilter(int col, std::function<bool(const QVariant&)> predicate) {
    if (!m_store) return;

    int rowCount = m_endRow - m_startRow + 1;
    int words = (rowCount + 63) / 64;
    std::vector<uint64_t> colBitmap(words, 0);

    m_store->scanColumnValues(col, m_startRow, m_endRow,
        [&](int row, const QVariant& val) {
            if (predicate(val)) setBit(colBitmap, row);
        });

    bool found = false;
    for (auto& cf : m_columnFilters) {
        if (cf.col == col) {
            cf.bitmap = std::move(colBitmap);
            found = true;
            break;
        }
    }
    if (!found) {
        m_columnFilters.push_back({col, std::move(colBitmap)});
    }

    rebuildCombinedBitmap();
}

// ============================================================================
// Clear filters
// ============================================================================

void FilterEngine::clearFilter(int col) {
    m_columnFilters.erase(
        std::remove_if(m_columnFilters.begin(), m_columnFilters.end(),
            [col](const ColumnFilter& cf) { return cf.col == col; }),
        m_columnFilters.end());

    rebuildCombinedBitmap();
}

void FilterEngine::clearAllFilters() {
    m_columnFilters.clear();
    m_isFiltered = false;
    initBitmapAllPass();
    rebuildFilteredMapping();
}
