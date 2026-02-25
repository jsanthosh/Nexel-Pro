#ifndef DOCUMENTTHEME_H
#define DOCUMENTTHEME_H

#include <QString>
#include <QColor>
#include <array>
#include <vector>

// The 10 standard theme color slots (matches Excel/OOXML)
enum class ThemeColorIndex : int {
    Dark1    = 0,   // Usually black (window text)
    Light1   = 1,   // Usually white (window background)
    Dark2    = 2,   // Dark accent for headings
    Light2   = 3,   // Light accent for backgrounds
    Accent1  = 4,
    Accent2  = 5,
    Accent3  = 6,
    Accent4  = 7,
    Accent5  = 8,
    Accent6  = 9
};

// A document-level color theme (10 base colors + tint/shade variants)
//
// Color string convention used in CellStyle.foregroundColor / backgroundColor:
//   Absolute: "#RRGGBB"                  (e.g. "#FF0000")
//   Theme:    "theme:<index>:<tint>"      (e.g. "theme:4:0.4" = Accent1, 40% lighter)
//
// Tint factors:
//   +0.8 = 80% lighter, +0.6 = 60% lighter, +0.4 = 40% lighter
//   -0.25 = 25% darker, -0.5 = 50% darker
//    0.0  = base color
struct DocumentTheme {
    QString id;           // e.g. "office", "green"
    QString displayName;  // e.g. "Office", "Green"
    std::array<QColor, 10> colors;  // Indexed by ThemeColorIndex

    // Compute a tinted/shaded variant of a theme color (Excel-compatible HSL algorithm).
    // tintFactor: -1.0 (full black) to +1.0 (full white), 0.0 = unchanged
    static QColor applyTint(const QColor& base, double tintFactor);

    // Resolve a theme color string to a QColor.
    // Returns the resolved color, or invalid QColor if not a theme string.
    QColor resolveColor(const QString& colorStr) const;

    // Build a theme color string: "theme:<index>:<tint>"
    static QString makeThemeColorStr(int themeIndex, double tint = 0.0);

    // Check if a color string is a theme reference
    static bool isThemeColor(const QString& colorStr);

    // Resolve any color string (theme or absolute) to a QColor
    QColor resolveAnyColor(const QString& colorStr) const;
};

// Registry of predefined document themes
std::vector<DocumentTheme> getBuiltinDocumentThemes();

// The default theme (Office)
const DocumentTheme& defaultDocumentTheme();

// Standard tint factors for the 6-row palette
static constexpr double kThemeTints[] = { 0.0, 0.8, 0.6, 0.4, -0.25, -0.5 };
static constexpr int kThemeTintCount = 6;

// Theme color display names
QString themeColorName(int index, double tint);

#endif // DOCUMENTTHEME_H
