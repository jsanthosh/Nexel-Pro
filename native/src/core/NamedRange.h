#ifndef NAMEDRANGE_H
#define NAMEDRANGE_H

#include <QString>
#include "CellRange.h"

struct NamedRange {
    QString name;
    CellRange range;
    int sheetIndex = -1; // -1 = global (workbook-level)
    bool isGlobal = true;
};

#endif // NAMEDRANGE_H
