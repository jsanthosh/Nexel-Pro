#ifndef STYLETABLE_H
#define STYLETABLE_H

#include <vector>
#include <unordered_map>
#include <shared_mutex>
#include <functional>
#include "Cell.h"  // for CellStyle

// ============================================================================
// StyleTable — Global deduplication of cell styles
// ============================================================================
// Typical spreadsheets have <200 unique style combinations even with 20M cells.
// Instead of storing ~281 bytes per cell, each cell stores a 2-byte index.
// Memory: 200 styles x 281 bytes = 56 KB total vs millions of CellStyle copies.
class StyleTable {
public:
    static StyleTable& instance();

    // Intern a style: returns a uint16_t index.
    // If an identical style already exists, returns its existing index.
    // Index 0 is always the default style.
    uint16_t intern(const CellStyle& style);

    // Get style by index. Returns default style for index 0.
    const CellStyle& get(uint16_t index) const;

    // Modify a style and get a new index (styles are immutable once interned).
    // Takes the old index, applies a modifier function, and returns the new index.
    uint16_t modify(uint16_t currentIndex,
                    const std::function<void(CellStyle&)>& modifier);

    // Number of unique styles
    size_t count() const;

    // Clear all styles (reset to just the default)
    void clear();

    // Serialization helpers
    const std::vector<CellStyle>& allStyles() const { return m_styles; }
    uint16_t addStyle(const CellStyle& style); // direct add without dedup (for deserialization)

private:
    StyleTable();

    // Hash a CellStyle for deduplication lookup
    static size_t hashStyle(const CellStyle& style);

    // Compare two styles for equality
    static bool stylesEqual(const CellStyle& a, const CellStyle& b);

    mutable std::shared_mutex m_mutex;
    std::vector<CellStyle> m_styles;            // index → style (index 0 = default)
    std::unordered_map<size_t, std::vector<uint16_t>> m_hashToIndices; // hash → candidate indices
};

#endif // STYLETABLE_H
