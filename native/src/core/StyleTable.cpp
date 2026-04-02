#include "StyleTable.h"
#include <QHash>
#include <mutex>

StyleTable& StyleTable::instance() {
    static StyleTable s;
    return s;
}

StyleTable::StyleTable() {
    // Index 0 = default style (always present)
    m_styles.emplace_back(CellStyle());
}

size_t StyleTable::hashStyle(const CellStyle& s) {
    // Combine hashes of all style fields
    size_t h = 0;
    auto combine = [&h](size_t v) {
        h ^= v + 0x9e3779b9 + (h << 6) + (h >> 2);
    };

    combine(qHash(s.fontName));
    combine(qHash(s.foregroundColor));
    combine(qHash(s.backgroundColor));
    combine(qHash(s.numberFormat));
    combine(qHash(s.currencyCode));
    combine(qHash(s.dateFormatId));
    combine(std::hash<int>()(s.fontSize));
    combine(std::hash<int>()(s.columnWidth));
    combine(std::hash<int>()(s.rowHeight));
    combine(std::hash<int>()(s.indentLevel));
    combine(std::hash<int>()(s.textRotation));
    combine(std::hash<uint8_t>()(static_cast<uint8_t>(s.hAlign)));
    combine(std::hash<uint8_t>()(static_cast<uint8_t>(s.vAlign)));
    combine(std::hash<uint8_t>()(static_cast<uint8_t>(s.textOverflow)));
    combine(std::hash<uint8_t>()(s.decimalPlaces));

    // Pack boolean flags into a single byte for hashing
    uint8_t flags = (s.bold) | (s.italic << 1) | (s.underline << 2) |
                    (s.strikethrough << 3) | (s.useThousandsSeparator << 4) |
                    (s.locked << 5) | (s.hidden << 6);
    combine(std::hash<uint8_t>()(flags));

    // Borders
    auto hashBorder = [&combine](const BorderStyle& b) {
        combine(qHash(b.color));
        uint8_t bp = (b.width) | (b.penStyle << 2) | (b.enabled << 4);
        combine(std::hash<uint8_t>()(bp));
    };
    hashBorder(s.borderTop);
    hashBorder(s.borderBottom);
    hashBorder(s.borderLeft);
    hashBorder(s.borderRight);

    return h;
}

bool StyleTable::stylesEqual(const CellStyle& a, const CellStyle& b) {
    return a.fontName == b.fontName &&
           a.foregroundColor == b.foregroundColor &&
           a.backgroundColor == b.backgroundColor &&
           a.numberFormat == b.numberFormat &&
           a.currencyCode == b.currencyCode &&
           a.dateFormatId == b.dateFormatId &&
           a.fontSize == b.fontSize &&
           a.columnWidth == b.columnWidth &&
           a.rowHeight == b.rowHeight &&
           a.indentLevel == b.indentLevel &&
           a.textRotation == b.textRotation &&
           a.hAlign == b.hAlign &&
           a.vAlign == b.vAlign &&
           a.textOverflow == b.textOverflow &&
           a.decimalPlaces == b.decimalPlaces &&
           a.bold == b.bold &&
           a.italic == b.italic &&
           a.underline == b.underline &&
           a.strikethrough == b.strikethrough &&
           a.useThousandsSeparator == b.useThousandsSeparator &&
           a.locked == b.locked &&
           a.hidden == b.hidden &&
           a.borderTop.color == b.borderTop.color &&
           a.borderTop.width == b.borderTop.width &&
           a.borderTop.penStyle == b.borderTop.penStyle &&
           a.borderTop.enabled == b.borderTop.enabled &&
           a.borderBottom.color == b.borderBottom.color &&
           a.borderBottom.width == b.borderBottom.width &&
           a.borderBottom.penStyle == b.borderBottom.penStyle &&
           a.borderBottom.enabled == b.borderBottom.enabled &&
           a.borderLeft.color == b.borderLeft.color &&
           a.borderLeft.width == b.borderLeft.width &&
           a.borderLeft.penStyle == b.borderLeft.penStyle &&
           a.borderLeft.enabled == b.borderLeft.enabled &&
           a.borderRight.color == b.borderRight.color &&
           a.borderRight.width == b.borderRight.width &&
           a.borderRight.penStyle == b.borderRight.penStyle &&
           a.borderRight.enabled == b.borderRight.enabled;
}

uint16_t StyleTable::intern(const CellStyle& style) {
    size_t h = hashStyle(style);

    // Fast path: read lock to check if style already exists
    {
        std::shared_lock readLock(m_mutex);
        auto it = m_hashToIndices.find(h);
        if (it != m_hashToIndices.end()) {
            for (uint16_t idx : it->second) {
                if (stylesEqual(m_styles[idx], style)) {
                    return idx;
                }
            }
        }
    }

    // Slow path: write lock to add new style
    std::unique_lock writeLock(m_mutex);

    // Double-check after acquiring write lock
    auto it = m_hashToIndices.find(h);
    if (it != m_hashToIndices.end()) {
        for (uint16_t idx : it->second) {
            if (stylesEqual(m_styles[idx], style)) {
                return idx;
            }
        }
    }

    // Check for overflow (uint16_t max = 65535)
    if (m_styles.size() >= 65535) {
        // Extremely unlikely: fall back to default style
        return 0;
    }

    uint16_t newIdx = static_cast<uint16_t>(m_styles.size());
    m_styles.push_back(style);
    m_hashToIndices[h].push_back(newIdx);
    return newIdx;
}

const CellStyle& StyleTable::get(uint16_t index) const {
    std::shared_lock readLock(m_mutex);
    if (index >= m_styles.size()) return m_styles[0]; // default
    return m_styles[index];
}

uint16_t StyleTable::modify(uint16_t currentIndex,
                             const std::function<void(CellStyle&)>& modifier) {
    CellStyle modified = get(currentIndex); // copy
    modifier(modified);
    return intern(modified);
}

size_t StyleTable::count() const {
    std::shared_lock readLock(m_mutex);
    return m_styles.size();
}

void StyleTable::clear() {
    std::unique_lock writeLock(m_mutex);
    m_styles.clear();
    m_hashToIndices.clear();
    m_styles.emplace_back(CellStyle()); // re-add default at index 0
}

uint16_t StyleTable::addStyle(const CellStyle& style) {
    std::unique_lock writeLock(m_mutex);
    uint16_t idx = static_cast<uint16_t>(m_styles.size());
    m_styles.push_back(style);
    size_t h = hashStyle(style);
    m_hashToIndices[h].push_back(idx);
    return idx;
}
