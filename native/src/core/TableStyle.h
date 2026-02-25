#ifndef TABLESTYLE_H
#define TABLESTYLE_H

#include <QString>
#include <QColor>
#include <QStringList>
#include <vector>
#include "CellRange.h"
#include "DocumentTheme.h"

struct TableTheme {
    QString name;
    QColor headerBg;
    QColor headerFg;
    QColor bandedRow1;
    QColor bandedRow2;
    QColor borderColor;
};

struct SpreadsheetTable {
    CellRange range;
    QString name;
    TableTheme theme;
    bool hasHeaderRow = true;
    bool bandedRows = true;
    QStringList columnNames;
};

// Predefined table themes (matching AdvancedSpreadsheet)
inline std::vector<TableTheme> getBuiltinTableThemes() {
    return {
        {"Ocean Blue",  QColor("#4472C4"), QColor("#FFFFFF"), QColor("#D9E2F3"), QColor("#FFFFFF"), QColor("#8FAADC")},
        {"Teal",        QColor("#2B9E8F"), QColor("#FFFFFF"), QColor("#D4EFED"), QColor("#FFFFFF"), QColor("#6DC4B9")},
        {"Indigo",      QColor("#4F46E5"), QColor("#FFFFFF"), QColor("#E0E7FF"), QColor("#FFFFFF"), QColor("#9B93F0")},
        {"Purple",      QColor("#7C3AED"), QColor("#FFFFFF"), QColor("#EDE9FE"), QColor("#FFFFFF"), QColor("#B794F4")},
        {"Rose",        QColor("#E11D48"), QColor("#FFFFFF"), QColor("#FFE4E6"), QColor("#FFFFFF"), QColor("#F06292")},
        {"Sunset",      QColor("#EA580C"), QColor("#FFFFFF"), QColor("#FED7AA"), QColor("#FFFFFF"), QColor("#F4A460")},
        {"Forest",      QColor("#16A34A"), QColor("#FFFFFF"), QColor("#DCFCE7"), QColor("#FFFFFF"), QColor("#6FCF97")},
        {"Slate",       QColor("#475569"), QColor("#FFFFFF"), QColor("#E2E8F0"), QColor("#FFFFFF"), QColor("#94A3B8")},
        {"Amber",       QColor("#D97706"), QColor("#FFFFFF"), QColor("#FEF3C7"), QColor("#FFFFFF"), QColor("#F0C674")},
        {"Cyan",        QColor("#0891B2"), QColor("#FFFFFF"), QColor("#CFFAFE"), QColor("#FFFFFF"), QColor("#67D3E8")},
        {"Blossom",     QColor("#DB2777"), QColor("#FFFFFF"), QColor("#FCE7F3"), QColor("#FFFFFF"), QColor("#F472B6")},
        {"Minimal",     QColor("#F3F4F6"), QColor("#111111"), QColor("#F9FAFB"), QColor("#FFFFFF"), QColor("#D1D5DB")},
    };
}

// Generate table themes from a document theme's 6 accent colors (2 per accent: light + medium)
inline std::vector<TableTheme> generateTableThemes(const DocumentTheme& theme) {
    static const char* kAccentNames[] = {
        "Accent 1", "Accent 2", "Accent 3", "Accent 4", "Accent 5", "Accent 6"
    };

    std::vector<TableTheme> themes;
    themes.reserve(12);

    for (int i = 0; i < 6; ++i) {
        QColor accent = theme.colors[4 + i]; // Accent1 = index 4

        // Light variant: colored header, very light banded rows
        QColor lightBand = DocumentTheme::applyTint(accent, 0.8);   // 80% lighter
        QColor border    = DocumentTheme::applyTint(accent, 0.4);   // 40% lighter
        themes.push_back({
            QString("%1 Light").arg(kAccentNames[i]),
            accent, QColor("#FFFFFF"), lightBand, QColor("#FFFFFF"), border
        });

        // Medium variant: colored header, medium banded rows
        QColor medBand = DocumentTheme::applyTint(accent, 0.6);     // 60% lighter
        themes.push_back({
            QString("%1 Medium").arg(kAccentNames[i]),
            accent, QColor("#FFFFFF"), medBand, QColor("#FFFFFF"), accent
        });
    }

    return themes;
}

#endif // TABLESTYLE_H
