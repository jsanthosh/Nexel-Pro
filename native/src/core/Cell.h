#ifndef CELL_H
#define CELL_H

#include <QString>
#include <QVariant>
#include <vector>
#include <memory>
#include <unordered_map>
#include <cstdint>
#include "CellRange.h"

enum class CellType {
    Empty,
    Text,
    Number,
    Formula,
    Date,
    Boolean,
    Error
};

enum class HorizontalAlignment : uint8_t {
    General,  // Auto: numbers right-align, text left-align
    Left,
    Center,
    Right
};

enum class VerticalAlignment : uint8_t {
    Top,
    Middle,
    Bottom
};

enum class TextOverflowMode : uint8_t {
    Overflow,     // Default: text spills into empty neighbor cells
    Wrap,         // Text wraps within the cell
    ShrinkToFit   // Font shrinks to fit within the cell
};

struct BorderStyle {
    QString color = "#000000";
    uint8_t width : 2;    // 1=thin, 2=medium, 3=thick (fits in 2 bits)
    uint8_t penStyle : 2; // 0=solid, 1=dashed, 2=dotted (fits in 2 bits)
    uint8_t enabled : 1;  // boolean flag as single bit

    BorderStyle() : width(1), penStyle(0), enabled(0) {}
};

// Color string convention for foregroundColor / backgroundColor / BorderStyle.color:
//   Absolute: "#RRGGBB"              (e.g. "#FF0000")
//   Theme:    "theme:<index>:<tint>"  (e.g. "theme:4:0.4" = Accent1, 40% lighter)
// See DocumentTheme.h for details.
struct CellStyle {
    // --- Pointer-aligned members first (8 bytes each on 64-bit) ---
    QString fontName = "Arial";
    QString foregroundColor = "#000000";
    QString backgroundColor = "#FFFFFF";
    QString numberFormat = "General";
    QString currencyCode = "USD";
    QString dateFormatId = "mm/dd/yyyy";

    // --- Borders (contain QString + bitfield each) ---
    BorderStyle borderTop;
    BorderStyle borderBottom;
    BorderStyle borderLeft;
    BorderStyle borderRight;

    // --- 4-byte members grouped together ---
    int fontSize = 11;
    int columnWidth = 80;
    int rowHeight = 22;
    int indentLevel = 0;
    int textRotation = 0;    // degrees: 0, 45, 90, -45, -90, 270=vertical stack

    // --- Small enum members (1 byte each with uint8_t underlying type) ---
    HorizontalAlignment hAlign = HorizontalAlignment::General;
    VerticalAlignment vAlign = VerticalAlignment::Middle;
    TextOverflowMode textOverflow = TextOverflowMode::Overflow;
    uint8_t decimalPlaces = 2;

    // Boolean flags packed as bitfields (1 byte total instead of 5 bytes)
    uint8_t bold : 1;
    uint8_t italic : 1;
    uint8_t underline : 1;
    uint8_t strikethrough : 1;
    uint8_t useThousandsSeparator : 1;
    uint8_t _reserved : 3;

    CellStyle()
        : bold(0), italic(0), underline(0), strikethrough(0),
          useThousandsSeparator(0), _reserved(0) {}
};

class Cell {
public:
    Cell();
    ~Cell() = default;

    // Value management
    void setValue(const QVariant& value);
    void setFormula(const QString& formula);
    QVariant getValue() const;
    QString getFormula() const;
    CellType getType() const;

    // Styling — lazy: default style shared across all cells, custom allocated on demand
    void setStyle(const CellStyle& style);
    const CellStyle& getStyle() const;
    bool hasCustomStyle() const { return m_customStyle != nullptr; }

    // Computed value (for formulas)
    void setComputedValue(const QVariant& value);
    QVariant getComputedValue() const;

    // State
    bool isDirty() const;
    void setDirty(bool dirty);

    bool hasError() const;
    void setError(const QString& error);
    QString getError() const;

    // Comments/Notes
    QString getComment() const;
    void setComment(const QString& comment);
    bool hasComment() const;

    // Hyperlinks
    QString getHyperlink() const { return m_hyperlink; }
    void setHyperlink(const QString& url) { m_hyperlink = url; }
    bool hasHyperlink() const { return !m_hyperlink.isEmpty(); }

    // Spill tracking (dynamic arrays)
    CellAddress getSpillParent() const { return m_spillParent; }
    void setSpillParent(const CellAddress& parent) { m_spillParent = parent; }
    bool isSpillCell() const { return m_spillParent.row >= 0; }
    void clearSpillParent() { m_spillParent = CellAddress(-1, -1); }

    // Fast bulk import setters — skip type detection, caller must know the type
    void setValueDirect(double num) { m_value = num; m_type = CellType::Number; }
    void setValueDirect(const QString& text) { m_value = text; m_type = CellType::Text; }

    // Utilities
    QString toString() const;
    void clear();

    // Shared default style (single allocation, reused by all cells)
    static const CellStyle& defaultStyle();

private:
    // --- Pointer-sized / large members first to minimize padding ---
    QVariant m_value;
    QVariant m_computedValue;
    QString m_formula;
    std::unique_ptr<CellStyle> m_customStyle; // null = default style
    QString m_error;
    QString m_comment;
    QString m_hyperlink;
    // --- Smaller members packed at end ---
    CellAddress m_spillParent{-1, -1}; // if spill cell, points to parent formula cell
    CellType m_type;
    bool m_dirty;
};

#endif // CELL_H
