#include "CellRange.h"
#include <QRegularExpression>
#include <algorithm>

QString CellAddress::toString() const {
    // Convert column index to letter (0 -> A, 1 -> B, etc.)
    QString colStr;
    int c = col;
    while (c >= 0) {
        colStr = QChar('A' + (c % 26)) + colStr;
        c = c / 26 - 1;
    }
    return colStr + QString::number(row + 1);
}

CellAddress CellAddress::fromString(const QString& str) {
    // Parse "A1", "$A$1", "$A1", "A$1" formats
    static constexpr int MAX_COL = 16384;  // XFD (Excel max)
    static constexpr int MAX_ROW = 100000000; // 100M rows
    int col = 0;
    int row = 0;
    int i = 0;

    // Skip optional $ before column
    while (i < str.length() && str[i] == '$') i++;

    // Parse column letters with overflow protection
    while (i < str.length() && str[i].isLetter()) {
        col = col * 26 + (str[i].toUpper().toLatin1() - 'A' + 1);
        if (col > MAX_COL + 1) {
            return CellAddress(-1, -1); // Invalid: column overflow
        }
        i++;
    }
    col--; // Convert to 0-based

    // Skip optional $ before row
    while (i < str.length() && str[i] == '$') i++;

    // Parse row number with overflow protection
    while (i < str.length() && str[i].isDigit()) {
        row = row * 10 + str[i].toLatin1() - '0';
        if (row > MAX_ROW) {
            return CellAddress(-1, -1); // Invalid: row overflow
        }
        i++;
    }
    row--; // Convert to 0-based

    return CellAddress(row, col);
}

CellRange::CellRange() : m_start(0, 0), m_end(0, 0) {
}

CellRange::CellRange(const CellAddress& start, const CellAddress& end)
    : m_start(start), m_end(end) {
    normalize();
}

CellRange::CellRange(int startRow, int startCol, int endRow, int endCol)
    : m_start(startRow, startCol), m_end(endRow, endCol) {
    normalize();
}

CellRange::CellRange(const QString& rangeStr) {
    // Handle "A1:B10" or just "A1"
    if (rangeStr.contains(':')) {
        QStringList parts = rangeStr.split(':');
        m_start = CellAddress::fromString(parts[0]);
        m_end = CellAddress::fromString(parts[1]);
    } else {
        m_start = CellAddress::fromString(rangeStr);
        m_end = m_start;
    }
    normalize();
}

void CellRange::normalize() {
    if (m_end < m_start) {
        std::swap(m_start, m_end);
    }
}

CellAddress CellRange::getStart() const {
    return m_start;
}

CellAddress CellRange::getEnd() const {
    return m_end;
}

int CellRange::getRowCount() const {
    return m_end.row - m_start.row + 1;
}

int CellRange::getColumnCount() const {
    return m_end.col - m_start.col + 1;
}

std::vector<CellAddress> CellRange::getCells() const {
    std::vector<CellAddress> cells;
    for (int r = m_start.row; r <= m_end.row; ++r) {
        for (int c = m_start.col; c <= m_end.col; ++c) {
            cells.emplace_back(r, c);
        }
    }
    return cells;
}

bool CellRange::contains(const CellAddress& addr) const {
    return addr.row >= m_start.row && addr.row <= m_end.row &&
           addr.col >= m_start.col && addr.col <= m_end.col;
}

bool CellRange::contains(int row, int col) const {
    return contains(CellAddress(row, col));
}

bool CellRange::intersects(const CellRange& other) const {
    return !(m_end.row < other.m_start.row || m_start.row > other.m_end.row ||
             m_end.col < other.m_start.col || m_start.col > other.m_end.col);
}

QString CellRange::toString() const {
    if (isSingleCell()) {
        return m_start.toString();
    }
    return m_start.toString() + ":" + m_end.toString();
}

bool CellRange::isValid() const {
    return m_start.row >= 0 && m_start.col >= 0 && m_end.row >= 0 && m_end.col >= 0;
}

bool CellRange::isSingleCell() const {
    return m_start == m_end;
}

bool CellRange::isSingleRow() const {
    return m_start.row == m_end.row;
}

bool CellRange::isSingleColumn() const {
    return m_start.col == m_end.col;
}
