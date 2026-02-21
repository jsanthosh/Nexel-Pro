#ifndef CELL_H
#define CELL_H

#include <QString>
#include <QVariant>
#include <vector>
#include <memory>
#include <unordered_map>

enum class CellType {
    Empty,
    Text,
    Number,
    Formula,
    Date,
    Boolean,
    Error
};

enum class HorizontalAlignment {
    General,  // Auto: numbers right-align, text left-align
    Left,
    Center,
    Right
};

enum class VerticalAlignment {
    Top,
    Middle,
    Bottom
};

struct BorderStyle {
    bool enabled = false;
    QString color = "#000000";
    int width = 1; // 1=thin, 2=medium, 3=thick
};

struct CellStyle {
    QString fontName = "Arial";
    int fontSize = 11;
    bool bold = false;
    bool italic = false;
    bool underline = false;
    bool strikethrough = false;
    QString foregroundColor = "#000000";
    QString backgroundColor = "#FFFFFF";
    HorizontalAlignment hAlign = HorizontalAlignment::General;
    VerticalAlignment vAlign = VerticalAlignment::Middle;
    QString numberFormat = "General";
    int decimalPlaces = 2;
    bool useThousandsSeparator = false;
    QString currencyCode = "USD";
    QString dateFormatId = "mm/dd/yyyy";
    int columnWidth = 80;
    int rowHeight = 22;
    // Borders
    BorderStyle borderTop;
    BorderStyle borderBottom;
    BorderStyle borderLeft;
    BorderStyle borderRight;
    // Indent
    int indentLevel = 0;
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

    // Utilities
    QString toString() const;
    void clear();

    // Shared default style (single allocation, reused by all cells)
    static const CellStyle& defaultStyle();

private:
    QVariant m_value;
    QString m_formula;
    QVariant m_computedValue;
    CellType m_type;
    std::unique_ptr<CellStyle> m_customStyle; // null = default style
    bool m_dirty;
    QString m_error;
};

#endif // CELL_H
