#ifndef CELLRANGE_H
#define CELLRANGE_H

#include <QString>
#include <QVariant>
#include <vector>

struct CellAddress {
    int row;
    int col;

    CellAddress(int r = 0, int c = 0) : row(r), col(c) {}

    bool operator==(const CellAddress& other) const {
        return row == other.row && col == other.col;
    }

    bool operator<(const CellAddress& other) const {
        if (row != other.row) return row < other.row;
        return col < other.col;
    }

    QString toString() const;
    static CellAddress fromString(const QString& str);
};

class CellRange {
public:
    CellRange();
    CellRange(const CellAddress& start, const CellAddress& end);
    CellRange(int startRow, int startCol, int endRow, int endCol);
    explicit CellRange(const QString& rangeStr); // "A1:B10"

    CellAddress getStart() const;
    CellAddress getEnd() const;
    int getRowCount() const;
    int getColumnCount() const;
    std::vector<CellAddress> getCells() const;

    bool contains(const CellAddress& addr) const;
    bool contains(int row, int col) const;
    bool intersects(const CellRange& other) const;

    QString toString() const;
    bool isValid() const;
    bool isSingleCell() const;
    bool isSingleRow() const;
    bool isSingleColumn() const;

private:
    CellAddress m_start;
    CellAddress m_end;

    void normalize();
};

Q_DECLARE_METATYPE(CellRange)

#endif // CELLRANGE_H
