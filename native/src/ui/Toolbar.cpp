#include "Toolbar.h"
#include "Theme.h"
#include <QAction>
#include <QToolButton>
#include <QSpinBox>
#include <QFontComboBox>
#include <QColorDialog>
#include <QLabel>
#include <QPainter>
#include <QPainterPath>
#include <QPixmap>
#include <QMenu>
#include <QPolygon>
#include <QGridLayout>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QWidgetAction>
#include <QTimer>
#include <QFrame>
#include <QPushButton>
#include <functional>
#include "../core/TableStyle.h"
#include "../core/NumberFormat.h"
#include "../core/DocumentTheme.h"
#include <QDate>
#include <QDateTime>

// ===== Color Palette Popup =====
// Custom paint widget for a single color swatch — crisp, no stylesheet blur
class ColorSwatch : public QWidget {
public:
    QColor color;
    bool selected = false;
    bool hovered = false;
    std::function<void()> onClick;

    ColorSwatch(const QColor& c, QWidget* parent) : QWidget(parent), color(c) {
        setFixedSize(18, 18);
        setCursor(Qt::PointingHandCursor);
        setAttribute(Qt::WA_Hover, true);
        setMouseTracking(true);
    }

protected:
    void paintEvent(QPaintEvent*) override {
        QPainter p(this);
        p.setRenderHint(QPainter::Antialiasing, false);
        QRect r = rect();

        // Fill
        p.fillRect(r, color);

        // Border
        if (selected) {
            p.setPen(QPen(ThemeManager::instance().currentTheme().accentDark, 2));
            p.drawRect(r.adjusted(0, 0, -1, -1));
        } else if (hovered) {
            p.setPen(QPen(QColor("#333333"), 2));
            p.drawRect(r.adjusted(0, 0, -1, -1));
        } else {
            // Subtle border only for light colors
            if (color.lightness() > 220) {
                p.setPen(QPen(QColor("#D0D0D0"), 1));
                p.drawRect(r.adjusted(0, 0, -1, -1));
            }
        }
    }

    void enterEvent(QEnterEvent*) override { hovered = true; update(); }
    void leaveEvent(QEvent*) override { hovered = false; update(); }
    void mousePressEvent(QMouseEvent*) override { if (onClick) onClick(); }
};

// Result from the color picker: resolved QColor for UI + color string for storage
struct ColorPickResult {
    QColor displayColor;   // Resolved QColor for immediate UI feedback
    QString colorString;   // "theme:4:0.4" or "#RRGGBB" — stored in CellStyle
    bool isValid = false;
};

static ColorPickResult showColorPalette(QWidget* parent, const QString& currentColorStr,
                                         const DocumentTheme& docTheme, const QString& title) {
    // Standard vivid colors row (matches Excel's standard colors)
    static const QColor standardColors[] = {
        QColor("#C00000"), QColor("#FF0000"), QColor("#FFC000"), QColor("#FFFF00"),
        QColor("#92D050"), QColor("#00B050"), QColor("#00B0F0"), QColor("#0070C0"),
        QColor("#002060"), QColor("#7030A0"),
    };
    // Grayscale row
    static const QColor grayscale[] = {
        QColor("#000000"), QColor("#1A1A1A"), QColor("#333333"), QColor("#4D4D4D"),
        QColor("#666666"), QColor("#808080"), QColor("#999999"), QColor("#B3B3B3"),
        QColor("#D9D9D9"), QColor("#FFFFFF"),
    };

    static const int COLS = 10;

    // Resolve current color for selected-swatch comparison
    QColor currentColor = DocumentTheme::isThemeColor(currentColorStr)
        ? docTheme.resolveColor(currentColorStr) : QColor(currentColorStr);

    QMenu menu(parent);
    menu.setStyleSheet(
        "QMenu { background: #FFFFFF; border: 1px solid #E0E0E0; padding: 0px; border-radius: 6px; }"
    );

    QWidget* container = new QWidget(&menu);
    QVBoxLayout* layout = new QVBoxLayout(container);
    layout->setContentsMargins(10, 8, 10, 8);
    layout->setSpacing(0);

    // Section: Theme Colors
    QLabel* themeLabel = new QLabel("Theme Colors", container);
    themeLabel->setStyleSheet("font: 11px 'Segoe UI', 'SF Pro Text', sans-serif; color: #666; padding-bottom: 4px;");
    layout->addWidget(themeLabel);

    ColorPickResult result;

    // Theme color grid: 6 rows × 10 columns
    QGridLayout* grid = new QGridLayout();
    grid->setSpacing(3);
    grid->setContentsMargins(0, 0, 0, 0);

    for (int r = 0; r < kThemeTintCount; ++r) {
        for (int c = 0; c < COLS; ++c) {
            QColor swatchColor = DocumentTheme::applyTint(docTheme.colors[c], kThemeTints[r]);
            ColorSwatch* swatch = new ColorSwatch(swatchColor, container);
            swatch->selected = currentColor.isValid() && (swatchColor == currentColor);
            swatch->setToolTip(themeColorName(c, kThemeTints[r]));
            swatch->onClick = [&result, &menu, c, r, swatchColor]() {
                result.displayColor = swatchColor;
                result.colorString = DocumentTheme::makeThemeColorStr(c, kThemeTints[r]);
                result.isValid = true;
                menu.close();
            };
            grid->addWidget(swatch, r, c);
        }
    }
    layout->addLayout(grid);

    // Separator
    layout->addSpacing(6);
    QFrame* sep1 = new QFrame(container);
    sep1->setFrameShape(QFrame::HLine);
    sep1->setStyleSheet("background: #E8E8E8; max-height: 1px;");
    layout->addWidget(sep1);
    layout->addSpacing(4);

    // Section: Standard Colors (vivid)
    QLabel* stdLabel = new QLabel("Standard Colors", container);
    stdLabel->setStyleSheet("font: 11px 'Segoe UI', 'SF Pro Text', sans-serif; color: #666; padding-bottom: 4px;");
    layout->addWidget(stdLabel);

    QHBoxLayout* stdRow = new QHBoxLayout();
    stdRow->setSpacing(3);
    stdRow->setContentsMargins(0, 0, 0, 0);
    for (int c = 0; c < COLS; ++c) {
        ColorSwatch* swatch = new ColorSwatch(standardColors[c], container);
        swatch->selected = currentColor.isValid() && (standardColors[c] == currentColor);
        swatch->setToolTip(standardColors[c].name().toUpper());
        swatch->onClick = [&result, &menu, c]() {
            result.displayColor = standardColors[c];
            result.colorString = standardColors[c].name();
            result.isValid = true;
            menu.close();
        };
        stdRow->addWidget(swatch);
    }
    layout->addLayout(stdRow);

    layout->addSpacing(4);

    // Grayscale row
    QHBoxLayout* grayRow = new QHBoxLayout();
    grayRow->setSpacing(3);
    grayRow->setContentsMargins(0, 0, 0, 0);
    for (int c = 0; c < COLS; ++c) {
        ColorSwatch* swatch = new ColorSwatch(grayscale[c], container);
        swatch->selected = currentColor.isValid() && (grayscale[c] == currentColor);
        swatch->setToolTip(grayscale[c].name().toUpper());
        swatch->onClick = [&result, &menu, c]() {
            result.displayColor = grayscale[c];
            result.colorString = grayscale[c].name();
            result.isValid = true;
            menu.close();
        };
        grayRow->addWidget(swatch);
    }
    layout->addLayout(grayRow);

    // Separator
    layout->addSpacing(6);
    QFrame* sep2 = new QFrame(container);
    sep2->setFrameShape(QFrame::HLine);
    sep2->setStyleSheet("background: #E8E8E8; max-height: 1px;");
    layout->addWidget(sep2);
    layout->addSpacing(4);

    // No Fill (for fill color only)
    if (title.contains("Fill")) {
        QPushButton* noFillBtn = new QPushButton("No Fill", container);
        noFillBtn->setFixedHeight(26);
        noFillBtn->setCursor(Qt::PointingHandCursor);
        noFillBtn->setStyleSheet(
            "QPushButton { background: transparent; border: none; font: 12px 'Segoe UI', 'SF Pro Text', sans-serif;"
            "  color: #444; text-align: left; padding-left: 2px; }"
            "QPushButton:hover { background: #F0F4F8; border-radius: 3px; }");
        QObject::connect(noFillBtn, &QPushButton::clicked, &menu, [&result, &menu]() {
            result.displayColor = QColor("#FFFFFF");
            result.colorString = "#FFFFFF";
            result.isValid = true;
            menu.close();
        });
        layout->addWidget(noFillBtn);
    }

    // Custom Color button
    QPushButton* customBtn = new QPushButton("Custom Color...", container);
    customBtn->setFixedHeight(26);
    customBtn->setCursor(Qt::PointingHandCursor);
    customBtn->setStyleSheet(
        "QPushButton { background: transparent; border: none; font: 12px 'Segoe UI', 'SF Pro Text', sans-serif;"
        "  color: #2980B9; text-align: left; padding-left: 2px; }"
        "QPushButton:hover { background: #F0F4F8; border-radius: 3px; }");
    QObject::connect(customBtn, &QPushButton::clicked, &menu, [&result, &menu, parent, currentColor, title]() {
        menu.close();
        QColor custom = QColorDialog::getColor(currentColor, parent, title);
        if (custom.isValid()) {
            result.displayColor = custom;
            result.colorString = custom.name();
            result.isValid = true;
        }
    });
    layout->addWidget(customBtn);

    QWidgetAction* wa = new QWidgetAction(&menu);
    wa->setDefaultWidget(container);
    menu.addAction(wa);

    menu.exec(QCursor::pos());
    return result;
}

// ===== Modern Icon Helpers (high-DPI aware, 20x20) =====

static QIcon createIcon(int size, std::function<void(QPainter&, int)> drawFunc) {
    qreal dpr = 2.0; // retina
    QPixmap pix(size * dpr, size * dpr);
    pix.setDevicePixelRatio(dpr);
    pix.fill(Qt::transparent);
    QPainter p(&pix);
    p.setRenderHint(QPainter::Antialiasing, true);
    drawFunc(p, size);
    p.end();
    return QIcon(pix);
}

static QIcon createNewIcon() {
    return createIcon(18, [](QPainter& p, int) {
        p.setPen(QPen(QColor("#555"), 1.2));
        p.setBrush(QColor("#FAFAFA"));
        p.drawRoundedRect(3, 1, 11, 15, 1.5, 1.5);
        // fold corner
        QPainterPath fold;
        fold.moveTo(10, 1);
        fold.lineTo(14, 5);
        fold.lineTo(10, 5);
        fold.closeSubpath();
        p.setBrush(QColor("#DDD"));
        p.drawPath(fold);
        // lines
        p.setPen(QPen(QColor("#AAA"), 0.8));
        p.drawLine(5, 7, 12, 7);
        p.drawLine(QPointF(5, 9.5), QPointF(11, 9.5));
        p.drawLine(5, 12, 9, 12);
    });
}

static QIcon createSaveIcon() {
    return createIcon(18, [](QPainter& p, int) {
        // Classic floppy disk icon with better detail
        p.setPen(QPen(QColor("#4A4A4A"), 1.0));
        p.setBrush(QColor("#5B9BD5"));
        p.drawRoundedRect(2, 2, 14, 14, 1.5, 1.5);
        // Metal slider area (top)
        p.setPen(Qt::NoPen);
        p.setBrush(QColor("#E8E8E8"));
        p.drawRect(5, 2, 8, 5);
        // Slider hole
        p.setBrush(QColor("#5B9BD5"));
        p.drawRect(QRectF(9, 2.5, 2.5, 4));
        // Label area (bottom)
        p.setBrush(QColor("#FAFAFA"));
        p.drawRoundedRect(4, 10, 10, 6, 1, 1);
        // Label lines
        p.setPen(QPen(QColor("#CCC"), 0.6));
        p.drawLine(QPointF(5.5, 12), QPointF(12.5, 12));
        p.drawLine(QPointF(5.5, 14), QPointF(10, 14));
    });
}

static QIcon createUndoRedoIcon(bool isUndo) {
    return createIcon(18, [isUndo](QPainter& p, int) {
        QColor color("#4A90D9");
        if (isUndo) {
            QPainterPath arc;
            arc.moveTo(4, 9);
            arc.cubicTo(4, 4, 9, 3, 14, 5);
            arc.cubicTo(16, 6, 16, 10, 14, 12);
            p.setPen(QPen(color, 1.8, Qt::SolidLine, Qt::RoundCap, Qt::RoundJoin));
            p.setBrush(Qt::NoBrush);
            p.drawPath(arc);
            p.setPen(Qt::NoPen);
            p.setBrush(color);
            QPolygonF arrow;
            arrow << QPointF(1.5, 9) << QPointF(5.5, 6.5) << QPointF(5.5, 11.5);
            p.drawPolygon(arrow);
        } else {
            QPainterPath arc;
            arc.moveTo(14, 9);
            arc.cubicTo(14, 4, 9, 3, 4, 5);
            arc.cubicTo(2, 6, 2, 10, 4, 12);
            p.setPen(QPen(color, 1.8, Qt::SolidLine, Qt::RoundCap, Qt::RoundJoin));
            p.setBrush(Qt::NoBrush);
            p.drawPath(arc);
            p.setPen(Qt::NoPen);
            p.setBrush(color);
            QPolygonF arrow;
            arrow << QPointF(16.5, 9) << QPointF(12.5, 6.5) << QPointF(12.5, 11.5);
            p.drawPolygon(arrow);
        }
    });
}

static QIcon createFormatPainterIcon() {
    return createIcon(18, [](QPainter& p, int) {
        // Modern flat paint roller icon
        p.setRenderHint(QPainter::Antialiasing, true);
        // Roller head
        p.setPen(QPen(QColor("#4A90D9"), 1.0));
        p.setBrush(QColor("#5BA3E6"));
        p.drawRoundedRect(3, 2, 12, 5, 2, 2);
        // Roller handle (vertical bar from roller to arm)
        p.setPen(Qt::NoPen);
        p.setBrush(QColor("#888"));
        p.drawRect(QRectF(8, 7, 2, 3));
        // Arm (horizontal)
        p.drawRect(QRectF(9, 9, 4, 2));
        // Handle (vertical grip)
        p.setBrush(QColor("#666"));
        p.drawRoundedRect(QRectF(11, 10, 2.5, 6), 1, 1);
    });
}

static QIcon createHAlignIcon(const QString& align) {
    return createIcon(16, [&align](QPainter& p, int) {
        p.setPen(QPen(QColor("#555"), 1.4, Qt::SolidLine, Qt::RoundCap));
        int widths[] = {11, 7, 10, 5};
        for (int i = 0; i < 4; ++i) {
            int y = 3 + i * 3;
            int w = widths[i];
            int x;
            if (align == "left") x = 2;
            else if (align == "center") x = (16 - w) / 2;
            else x = 14 - w;
            p.drawLine(x, y, x + w, y);
        }
    });
}

static QIcon createVAlignIcon(const QString& align) {
    return createIcon(16, [&align](QPainter& p, int) {
        // Cell border
        p.setPen(QPen(QColor("#BBB"), 0.8));
        p.setBrush(Qt::NoBrush);
        p.drawRoundedRect(1, 1, 13, 13, 1.5, 1.5);
        // Lines
        p.setPen(QPen(QColor("#555"), 1.4, Qt::SolidLine, Qt::RoundCap));
        int startY;
        if (align == "top") startY = 3;
        else if (align == "middle") startY = 5;
        else startY = 8;
        p.drawLine(3, startY, 12, startY);
        p.drawLine(3, startY + 3, 10, startY + 3);
    });
}

static QIcon createSortIcon(bool ascending) {
    return createIcon(16, [ascending](QPainter& p, int) {
        QColor accent("#4A90D9");
        p.setPen(QPen(accent, 1.5));
        p.setBrush(accent);
        if (ascending) {
            p.drawLine(11, 3, 11, 13);
            QPolygon ah;
            ah << QPoint(8, 10) << QPoint(11, 14) << QPoint(14, 10);
            p.drawPolygon(ah);
        } else {
            p.drawLine(11, 3, 11, 13);
            QPolygon ah;
            ah << QPoint(8, 6) << QPoint(11, 2) << QPoint(14, 6);
            p.drawPolygon(ah);
        }
        p.setPen(QColor("#555"));
        p.setBrush(Qt::NoBrush);
        QFont f("Arial", 6, QFont::Bold);
        p.setFont(f);
        if (ascending) {
            p.drawText(QRect(0, 0, 9, 9), Qt::AlignCenter, "A");
            p.drawText(QRect(0, 7, 9, 9), Qt::AlignCenter, "Z");
        } else {
            p.drawText(QRect(0, 0, 9, 9), Qt::AlignCenter, "Z");
            p.drawText(QRect(0, 7, 9, 9), Qt::AlignCenter, "A");
        }
    });
}

static QIcon createFilterIcon() {
    return createIcon(18, [](QPainter& p, int) {
        // Funnel shape like Excel
        p.setPen(QPen(QColor("#4A90D9"), 1.4, Qt::SolidLine, Qt::RoundCap, Qt::RoundJoin));
        p.setBrush(QColor("#4A90D9").lighter(180));
        QPolygonF funnel;
        funnel << QPointF(2, 3) << QPointF(16, 3)
               << QPointF(11, 9) << QPointF(11, 14)
               << QPointF(7, 14) << QPointF(7, 9);
        funnel << QPointF(2, 3);
        p.drawPolygon(funnel);
    });
}

static QIcon createTableIcon() {
    return createIcon(18, [](QPainter& p, int) {
        p.setPen(QPen(QColor("#555"), 0.8));
        // Header
        p.setBrush(QColor("#4A90D9"));
        p.drawRoundedRect(2, 2, 14, 4, 1.5, 1.5);
        // Rows
        p.setBrush(QColor("#E8F0FE"));
        p.drawRect(2, 6, 14, 4);
        p.setBrush(QColor("#FAFAFA"));
        p.drawRect(2, 10, 14, 4);
        // Grid lines
        p.setPen(QPen(QColor("#BBB"), 0.5));
        p.drawLine(9, 6, 9, 14);
        // Border
        p.setPen(QPen(QColor("#4A90D9"), 0.8));
        p.setBrush(Qt::NoBrush);
        p.drawRoundedRect(2, 2, 14, 12, 1.5, 1.5);
    });
}

static QIcon createCondFmtIcon() {
    return createIcon(18, [](QPainter& p, int) {
        p.setPen(Qt::NoPen);
        p.setBrush(QColor("#4CAF50"));
        p.drawRoundedRect(2, 2, 14, 4, 2, 2);
        p.setBrush(QColor("#FF9800"));
        p.drawRoundedRect(2, 7, 14, 4, 2, 2);
        p.setBrush(QColor("#F44336"));
        p.drawRoundedRect(2, 12, 14, 4, 2, 2);
    });
}

static QIcon createValidationIcon() {
    return createIcon(18, [](QPainter& p, int) {
        p.setPen(QPen(QColor("#4CAF50"), 1.5));
        p.setBrush(Qt::NoBrush);
        p.drawEllipse(2, 2, 14, 14);
        p.setPen(QPen(QColor("#4CAF50"), 2.0, Qt::SolidLine, Qt::RoundCap, Qt::RoundJoin));
        p.drawLine(6, 9, 8, 12);
        p.drawLine(8, 12, 13, 5);
    });
}

static QIcon createFormatCellsIcon() {
    return createIcon(18, [](QPainter& p, int) {
        p.setPen(QPen(QColor("#555"), 1.0));
        p.setBrush(QColor("#FAFAFA"));
        p.drawRoundedRect(2, 2, 14, 14, 2, 2);
        // grid
        p.setPen(QPen(QColor("#CCC"), 0.6));
        p.drawLine(2, 7, 16, 7);
        p.drawLine(2, 12, 16, 12);
        p.drawLine(9, 2, 9, 16);
        // small colored corner
        p.setPen(Qt::NoPen);
        p.setBrush(QColor("#4A90D9"));
        p.drawRoundedRect(12, 12, 4, 4, 1, 1);
    });
}

static QIcon createBorderIcon() {
    return createIcon(18, [](QPainter& p, int) {
        // Grid with border emphasis
        p.setPen(QPen(QColor("#555"), 1.6));
        p.setBrush(Qt::NoBrush);
        p.drawRect(2, 2, 14, 14);
        // Inner grid (lighter)
        p.setPen(QPen(QColor("#BBB"), 0.6));
        p.drawLine(9, 2, 9, 16);
        p.drawLine(2, 9, 16, 9);
    });
}

static QIcon createMergeIcon() {
    return createIcon(18, [](QPainter& p, int) {
        // Two cells becoming one
        p.setPen(QPen(QColor("#555"), 1.0));
        p.setBrush(QColor("#E8F0FE"));
        p.drawRoundedRect(2, 4, 14, 10, 1.5, 1.5);
        // Arrows pointing inward
        p.setPen(QPen(QColor("#4A90D9"), 1.8, Qt::SolidLine, Qt::RoundCap));
        p.drawLine(4, 9, 7, 9);
        p.drawLine(14, 9, 11, 9);
        // Arrow heads
        p.setPen(Qt::NoPen);
        p.setBrush(QColor("#4A90D9"));
        QPolygonF leftArrow;
        leftArrow << QPointF(7, 7) << QPointF(7, 11) << QPointF(9, 9);
        p.drawPolygon(leftArrow);
        QPolygonF rightArrow;
        rightArrow << QPointF(11, 7) << QPointF(11, 11) << QPointF(9, 9);
        p.drawPolygon(rightArrow);
    });
}

static QIcon createChatIcon() {
    return createIcon(18, [](QPainter& p, int) {
        const QColor accent = ThemeManager::instance().currentTheme().accentDarker;
        // Speech bubble
        p.setPen(QPen(accent, 1.2));
        p.setBrush(accent.lighter(160));
        QPainterPath bubble;
        bubble.addRoundedRect(2, 2, 14, 10, 3, 3);
        // Tail
        bubble.moveTo(5, 12);
        bubble.lineTo(4, 16);
        bubble.lineTo(9, 12);
        p.drawPath(bubble);
        // Dots
        p.setPen(Qt::NoPen);
        p.setBrush(accent);
        p.drawEllipse(QPointF(6, 7), 1.2, 1.2);
        p.drawEllipse(QPointF(9, 7), 1.2, 1.2);
        p.drawEllipse(QPointF(12, 7), 1.2, 1.2);
    });
}

// Border menu item icons
static QIcon createBorderMenuIcon(const QString& type) {
    return createIcon(16, [&type](QPainter& p, int) {
        // Light cell grid background
        p.setPen(QPen(QColor("#CCC"), 0.5));
        p.setBrush(Qt::NoBrush);
        p.drawRect(1, 1, 13, 13);
        p.drawLine(7, 1, 7, 14);
        p.drawLine(1, 7, 14, 7);

        if (type == "all") {
            p.setPen(QPen(QColor("#333"), 1.4));
            p.drawRect(1, 1, 13, 13);
            p.drawLine(7, 1, 7, 14);
            p.drawLine(1, 7, 14, 7);
        } else if (type == "outside") {
            p.setPen(QPen(QColor("#333"), 1.6));
            p.drawRect(1, 1, 13, 13);
        } else if (type == "thick_outside") {
            p.setPen(QPen(QColor("#333"), 2.4));
            p.drawRect(1, 1, 13, 13);
        } else if (type == "bottom") {
            p.setPen(QPen(QColor("#333"), 1.6));
            p.drawLine(1, 14, 14, 14);
        } else if (type == "top") {
            p.setPen(QPen(QColor("#333"), 1.6));
            p.drawLine(1, 1, 14, 1);
        } else if (type == "left") {
            p.setPen(QPen(QColor("#333"), 1.6));
            p.drawLine(1, 1, 1, 14);
        } else if (type == "right") {
            p.setPen(QPen(QColor("#333"), 1.6));
            p.drawLine(14, 1, 14, 14);
        } else if (type == "inside_h") {
            p.setPen(QPen(QColor("#333"), 1.4));
            p.drawLine(1, 7, 14, 7);
        } else if (type == "inside_v") {
            p.setPen(QPen(QColor("#333"), 1.4));
            p.drawLine(7, 1, 7, 14);
        } else if (type == "inside") {
            p.setPen(QPen(QColor("#333"), 1.4));
            p.drawLine(7, 1, 7, 14);
            p.drawLine(1, 7, 14, 7);
        } else if (type == "none") {
            p.setPen(QPen(QColor("#C00000"), 1.4));
            p.drawLine(3, 3, 12, 12);
            p.drawLine(12, 3, 3, 12);
        }
    });
}

static QIcon createIndentIcon(bool increase) {
    return createIcon(16, [increase](QPainter& p, int) {
        p.setPen(QPen(QColor("#555"), 1.4, Qt::SolidLine, Qt::RoundCap));
        // Text lines
        int indent = increase ? 5 : 2;
        p.drawLine(indent, 3, 14, 3);
        p.drawLine(indent, 6, 12, 6);
        p.drawLine(indent, 9, 14, 9);
        p.drawLine(indent, 12, 11, 12);
        // Arrow
        p.setPen(QPen(QColor("#4A90D9"), 1.6, Qt::SolidLine, Qt::RoundCap));
        if (increase) {
            p.drawLine(1, 7, 4, 7);
            p.setPen(Qt::NoPen);
            p.setBrush(QColor("#4A90D9"));
            QPolygonF arr;
            arr << QPointF(4, 5) << QPointF(4, 9) << QPointF(6, 7);
            p.drawPolygon(arr);
        } else {
            p.drawLine(5, 7, 2, 7);
            p.setPen(Qt::NoPen);
            p.setBrush(QColor("#4A90D9"));
            QPolygonF arr;
            arr << QPointF(2, 5) << QPointF(2, 9) << QPointF(0, 7);
            p.drawPolygon(arr);
        }
    });
}

// ===== Modern Toolbar Style =====

static QString buildToolbarStyle() {
    const auto& t = ThemeManager::instance().currentTheme();
    return QString(R"(
    QToolBar {
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 %1, stop:0.04 %1,
            stop:0.041 #FAFBFC, stop:1 %2);
        border-bottom: 1px solid %3;
        spacing: 1px;
        padding: 5px 8px 4px 8px;
    }
    QToolButton {
        background: transparent;
        border: 1px solid transparent;
        border-radius: 4px;
        padding: 3px 6px;
        margin: 0px 1px;
        font-size: 12px;
        color: %4;
    }
    QToolButton:hover {
        background-color: %5;
        border-color: %3;
    }
    QToolButton:pressed {
        background-color: %3;
    }
    QToolButton:checked {
        background-color: %6;
        border-color: %7;
    }
    QFontComboBox {
        max-width: 180px;
        min-width: 140px;
        height: 26px;
        border: 1px solid %3;
        border-radius: 4px;
        padding: 1px 4px 1px 6px;
        background: white;
        font-size: 12px;
        color: %4;
    }
    QFontComboBox:focus {
        border-color: %7;
    }
    QFontComboBox QAbstractItemView {
        min-width: 200px;
    }
    QSpinBox {
        max-width: 50px;
        height: 26px;
        border: 1px solid %3;
        border-radius: 4px;
        padding: 1px 6px;
        background: white;
        font-size: 12px;
        color: %4;
    }
    QSpinBox:focus {
        border-color: %7;
    }
    QToolBar::separator {
        width: 1px;
        background-color: %8;
        margin: 4px 4px;
    }
    )").arg(
        t.toolbarAccentStripe.name(),    // %1
        t.toolbarBackground.name(),      // %2
        t.toolbarBorder.name(),          // %3
        t.toolbarButtonText.name(),      // %4
        t.toolbarButtonHover.name(),     // %5
        t.toolbarButtonChecked.name(),   // %6
        t.toolbarButtonCheckedBorder.name(), // %7
        t.toolbarSeparator.name()        // %8
    );
}

Toolbar::Toolbar(QWidget* parent)
    : QToolBar("Standard Toolbar", parent) {

    setMovable(false);
    setFloatable(false);
    setIconSize(QSize(18, 18));
    setStyleSheet(buildToolbarStyle());

    createActions();
}

void Toolbar::createActions() {
    // ===== File =====
    QToolButton* newBtn = new QToolButton(this);
    newBtn->setIcon(createNewIcon());
    newBtn->setToolTip("New Document (Ctrl+N)");
    newBtn->setFixedSize(30, 28);
    addWidget(newBtn);
    connect(newBtn, &QToolButton::clicked, this, &Toolbar::newDocument);

    m_saveBtn = new QToolButton(this);
    m_saveBtn->setIcon(createSaveIcon());
    m_saveBtn->setToolTip("Save Document (Ctrl+S)");
    m_saveBtn->setFixedSize(30, 28);
    m_saveBtn->setEnabled(false); // Disabled by default (no unsaved changes)
    addWidget(m_saveBtn);
    connect(m_saveBtn, &QToolButton::clicked, this, &Toolbar::saveDocument);

    addSeparator();

    // ===== Undo/Redo =====
    QToolButton* undoBtn = new QToolButton(this);
    undoBtn->setIcon(createUndoRedoIcon(true));
    undoBtn->setToolTip("Undo (Ctrl+Z)");
    undoBtn->setFixedSize(30, 28);
    addWidget(undoBtn);
    connect(undoBtn, &QToolButton::clicked, this, &Toolbar::undo);

    QToolButton* redoBtn = new QToolButton(this);
    redoBtn->setIcon(createUndoRedoIcon(false));
    redoBtn->setToolTip("Redo (Ctrl+Y)");
    redoBtn->setFixedSize(30, 28);
    addWidget(redoBtn);
    connect(redoBtn, &QToolButton::clicked, this, &Toolbar::redo);

    addSeparator();

    // ===== Format Painter =====
    QToolButton* formatPainterBtn = new QToolButton(this);
    formatPainterBtn->setIcon(createFormatPainterIcon());
    formatPainterBtn->setToolTip("Format Painter");
    formatPainterBtn->setCheckable(true);
    formatPainterBtn->setFixedSize(30, 28);
    addWidget(formatPainterBtn);
    connect(formatPainterBtn, &QToolButton::clicked, this, &Toolbar::formatPainterToggled);

    addSeparator();

    // ===== Font =====
    m_fontCombo = new QFontComboBox(this);
    addWidget(m_fontCombo);
    connect(m_fontCombo, &QFontComboBox::currentFontChanged, this, [this](const QFont& font) {
        emit fontFamilyChanged(font.family());
    });

    m_fontSizeSpinBox = new QSpinBox(this);
    m_fontSizeSpinBox->setRange(6, 72);
    m_fontSizeSpinBox->setValue(11);
    addWidget(m_fontSizeSpinBox);
    connect(m_fontSizeSpinBox, &QSpinBox::valueChanged, this, &Toolbar::fontSizeChanged);

    addSeparator();

    // ===== B I U S =====
    auto makeFmtBtn = [this](const QString& label, const QFont& font, const QString& tip,
                              const QKeySequence& shortcut, auto signal) {
        QToolButton* btn = new QToolButton(this);
        btn->setText(label);
        btn->setFont(font);
        btn->setToolTip(tip);
        if (!shortcut.isEmpty()) btn->setShortcut(shortcut);
        btn->setCheckable(true);
        btn->setFixedSize(28, 28);
        addWidget(btn);
        connect(btn, &QToolButton::clicked, this, signal);
        return btn;
    };

    QFont boldFont("Arial", 12, QFont::Bold);
    m_boldBtn = makeFmtBtn("B", boldFont, "Bold (Ctrl+B)", QKeySequence::Bold, [this]() { emit bold(); });

    QFont italicFont("Arial", 12);
    italicFont.setItalic(true);
    m_italicBtn = makeFmtBtn("I", italicFont, "Italic (Ctrl+I)", QKeySequence::Italic, [this]() { emit italic(); });

    QFont underlineFont("Arial", 12);
    underlineFont.setUnderline(true);
    m_underlineBtn = makeFmtBtn("U", underlineFont, "Underline (Ctrl+U)", QKeySequence::Underline, [this]() { emit underline(); });

    QFont strikeFont("Arial", 12);
    strikeFont.setStrikeOut(true);
    m_strikethroughBtn = makeFmtBtn("S", strikeFont, "Strikethrough", QKeySequence(), [this]() { emit strikethrough(); });

    addSeparator();

    // ===== Colors =====
    m_fgColorBtn = new QToolButton(this);
    m_fgColorBtn->setText("A");
    m_fgColorBtn->setFont(QFont("Arial", 12, QFont::Bold));
    m_fgColorBtn->setToolTip("Font Color");
    m_fgColorBtn->setFixedSize(28, 28);
    m_fgColorBtn->setStyleSheet(
        "QToolButton { color: #C00000; font-weight: bold; border-bottom: 3px solid #C00000; border-radius: 4px; }");
    addWidget(m_fgColorBtn);
    connect(m_fgColorBtn, &QToolButton::clicked, this, [this]() {
        const DocumentTheme& dt = m_docTheme ? *m_docTheme : defaultDocumentTheme();
        auto pick = showColorPalette(this, m_lastFgColorStr, dt, "Font Color");
        if (pick.isValid) {
            m_lastFgColor = pick.displayColor;
            m_lastFgColorStr = pick.colorString;
            m_fgColorBtn->setStyleSheet(
                QString("QToolButton { color: %1; font-weight: bold; border-bottom: 3px solid %1; border-radius: 4px; }")
                    .arg(pick.displayColor.name()));
            emit foregroundColorChanged(pick.colorString, pick.displayColor);
        }
    });

    m_bgColorBtn = new QToolButton(this);
    m_bgColorBtn->setToolTip("Fill Color");
    m_bgColorBtn->setFixedSize(28, 28);
    addWidget(m_bgColorBtn);
    updateBgColorIcon();
    connect(m_bgColorBtn, &QToolButton::clicked, this, [this]() {
        const DocumentTheme& dt = m_docTheme ? *m_docTheme : defaultDocumentTheme();
        auto pick = showColorPalette(this, m_lastBgColorStr, dt, "Fill Color");
        if (pick.isValid) {
            m_lastBgColor = pick.displayColor;
            m_lastBgColorStr = pick.colorString;
            updateBgColorIcon();
            emit backgroundColorChanged(pick.colorString, pick.displayColor);
        }
    });

}

// ===== Row 2: Secondary Toolbar =====

static QString buildToolbarStyleRow2() {
    const auto& t = ThemeManager::instance().currentTheme();
    return QString(R"(
    QToolBar {
        background: %1;
        border-bottom: 1px solid %2;
        spacing: 1px;
        padding: 2px 8px 2px 8px;
    }
    QToolButton {
        background: transparent;
        border: 1px solid transparent;
        border-radius: 4px;
        padding: 2px 5px;
        margin: 0px 1px;
        font-size: 12px;
        color: %3;
    }
    QToolButton:hover {
        background-color: %4;
        border-color: %2;
    }
    QToolButton:pressed {
        background-color: %2;
    }
    QToolButton:checked {
        background-color: %5;
        border-color: %6;
    }
    QToolButton::menu-indicator {
        subcontrol-position: right center;
        subcontrol-origin: padding;
        width: 10px;
        height: 10px;
    }
    QToolBar::separator {
        width: 1px;
        background-color: %7;
        margin: 3px 4px;
    }
    )").arg(
        t.toolbarBackground.name(),      // %1
        t.toolbarBorder.name(),          // %2
        t.toolbarButtonText.name(),      // %3
        t.toolbarButtonHover.name(),     // %4
        t.toolbarButtonChecked.name(),   // %5
        t.toolbarButtonCheckedBorder.name(), // %6
        t.toolbarSeparator.name()        // %7
    );
}

void Toolbar::setSaveEnabled(bool enabled) {
    if (m_saveBtn) m_saveBtn->setEnabled(enabled);
}

QToolBar* Toolbar::createSecondaryToolbar(QWidget* parent) {
    QToolBar* bar = new QToolBar("Format Toolbar", parent);
    bar->setMovable(false);
    bar->setFloatable(false);
    bar->setIconSize(QSize(16, 16));
    bar->setStyleSheet(buildToolbarStyleRow2());

    // ===== Horizontal Alignment =====
    m_alignLeftBtn = new QToolButton(bar);
    m_alignLeftBtn->setIcon(createHAlignIcon("left"));
    m_alignLeftBtn->setToolTip("Align Left");
    m_alignLeftBtn->setCheckable(true);
    m_alignLeftBtn->setFixedSize(26, 24);
    bar->addWidget(m_alignLeftBtn);

    m_alignCenterBtn = new QToolButton(bar);
    m_alignCenterBtn->setIcon(createHAlignIcon("center"));
    m_alignCenterBtn->setToolTip("Center");
    m_alignCenterBtn->setCheckable(true);
    m_alignCenterBtn->setFixedSize(26, 24);
    bar->addWidget(m_alignCenterBtn);

    m_alignRightBtn = new QToolButton(bar);
    m_alignRightBtn->setIcon(createHAlignIcon("right"));
    m_alignRightBtn->setToolTip("Align Right");
    m_alignRightBtn->setCheckable(true);
    m_alignRightBtn->setFixedSize(26, 24);
    bar->addWidget(m_alignRightBtn);

    auto uncheckHAligns = [this](QToolButton* except) {
        if (m_alignLeftBtn != except) m_alignLeftBtn->setChecked(false);
        if (m_alignCenterBtn != except) m_alignCenterBtn->setChecked(false);
        if (m_alignRightBtn != except) m_alignRightBtn->setChecked(false);
    };

    connect(m_alignLeftBtn, &QToolButton::clicked, this, [this, uncheckHAligns]() {
        uncheckHAligns(m_alignLeftBtn);
        emit hAlignChanged(HorizontalAlignment::Left);
    });
    connect(m_alignCenterBtn, &QToolButton::clicked, this, [this, uncheckHAligns]() {
        uncheckHAligns(m_alignCenterBtn);
        emit hAlignChanged(HorizontalAlignment::Center);
    });
    connect(m_alignRightBtn, &QToolButton::clicked, this, [this, uncheckHAligns]() {
        uncheckHAligns(m_alignRightBtn);
        emit hAlignChanged(HorizontalAlignment::Right);
    });

    bar->addSeparator();

    // ===== Vertical Alignment =====
    m_vAlignTopBtn = new QToolButton(bar);
    m_vAlignTopBtn->setIcon(createVAlignIcon("top"));
    m_vAlignTopBtn->setToolTip("Top Align");
    m_vAlignTopBtn->setCheckable(true);
    m_vAlignTopBtn->setFixedSize(26, 24);
    bar->addWidget(m_vAlignTopBtn);

    m_vAlignMiddleBtn = new QToolButton(bar);
    m_vAlignMiddleBtn->setIcon(createVAlignIcon("middle"));
    m_vAlignMiddleBtn->setToolTip("Middle Align");
    m_vAlignMiddleBtn->setCheckable(true);
    m_vAlignMiddleBtn->setChecked(true);
    m_vAlignMiddleBtn->setFixedSize(26, 24);
    bar->addWidget(m_vAlignMiddleBtn);

    m_vAlignBottomBtn = new QToolButton(bar);
    m_vAlignBottomBtn->setIcon(createVAlignIcon("bottom"));
    m_vAlignBottomBtn->setToolTip("Bottom Align");
    m_vAlignBottomBtn->setCheckable(true);
    m_vAlignBottomBtn->setFixedSize(26, 24);
    bar->addWidget(m_vAlignBottomBtn);

    auto uncheckVAligns = [this](QToolButton* except) {
        if (m_vAlignTopBtn != except) m_vAlignTopBtn->setChecked(false);
        if (m_vAlignMiddleBtn != except) m_vAlignMiddleBtn->setChecked(false);
        if (m_vAlignBottomBtn != except) m_vAlignBottomBtn->setChecked(false);
    };

    connect(m_vAlignTopBtn, &QToolButton::clicked, this, [this, uncheckVAligns]() {
        uncheckVAligns(m_vAlignTopBtn);
        emit vAlignChanged(VerticalAlignment::Top);
    });
    connect(m_vAlignMiddleBtn, &QToolButton::clicked, this, [this, uncheckVAligns]() {
        uncheckVAligns(m_vAlignMiddleBtn);
        emit vAlignChanged(VerticalAlignment::Middle);
    });
    connect(m_vAlignBottomBtn, &QToolButton::clicked, this, [this, uncheckVAligns]() {
        uncheckVAligns(m_vAlignBottomBtn);
        emit vAlignChanged(VerticalAlignment::Bottom);
    });

    bar->addSeparator();

    // ===== Indent =====
    QToolButton* indentIncBtn = new QToolButton(bar);
    indentIncBtn->setIcon(createIndentIcon(true));
    indentIncBtn->setToolTip("Increase Indent");
    indentIncBtn->setFixedSize(26, 24);
    bar->addWidget(indentIncBtn);
    connect(indentIncBtn, &QToolButton::clicked, this, &Toolbar::increaseIndent);

    QToolButton* indentDecBtn = new QToolButton(bar);
    indentDecBtn->setIcon(createIndentIcon(false));
    indentDecBtn->setToolTip("Decrease Indent");
    indentDecBtn->setFixedSize(26, 24);
    bar->addWidget(indentDecBtn);
    connect(indentDecBtn, &QToolButton::clicked, this, &Toolbar::decreaseIndent);

    // ===== Text Rotation =====
    QToolButton* rotateBtn = new QToolButton(bar);
    rotateBtn->setToolTip("Text Rotation");
    rotateBtn->setPopupMode(QToolButton::InstantPopup);
    rotateBtn->setFixedSize(38, 24);
    rotateBtn->setStyleSheet(
        "QToolButton { background: transparent; border: 1px solid transparent; border-radius: 4px; padding: 2px 4px; }"
        "QToolButton:hover { background-color: #E8ECF0; border-color: #D0D5DD; }"
    );
    // Create rotation icon: tilted "ab" text
    rotateBtn->setIcon(createIcon(16, [](QPainter& p, int) {
        p.setRenderHint(QPainter::Antialiasing, true);
        p.setPen(QPen(QColor("#555"), 1.2));
        QFont f = p.font();
        f.setPixelSize(9);
        f.setBold(true);
        p.setFont(f);
        p.translate(8, 8);
        p.rotate(-45);
        p.drawText(QRect(-10, -6, 20, 12), Qt::AlignCenter, "ab");
        p.rotate(45);
        p.translate(-8, -8);
    }));

    QMenu* rotateMenu = new QMenu(rotateBtn);
    rotateMenu->setStyleSheet(
        "QMenu { background: #FFFFFF; border: 1px solid #D0D5DD; border-radius: 6px; padding: 4px; }"
        "QMenu::item { padding: 5px 16px 5px 8px; border-radius: 4px; font-size: 12px; }"
        "QMenu::item:selected { background-color: #E8F0FE; }"
        "QMenu::separator { height: 1px; background: #E0E3E8; margin: 3px 8px; }"
    );
    rotateMenu->addAction("Angle Counterclockwise", this, [this]() { emit textRotationChanged(45); });
    rotateMenu->addAction("Angle Clockwise", this, [this]() { emit textRotationChanged(-45); });
    rotateMenu->addSeparator();
    rotateMenu->addAction("Vertical Text", this, [this]() { emit textRotationChanged(270); });
    rotateMenu->addAction("Rotate Text Up", this, [this]() { emit textRotationChanged(90); });
    rotateMenu->addAction("Rotate Text Down", this, [this]() { emit textRotationChanged(-90); });
    rotateMenu->addSeparator();
    rotateMenu->addAction("No Rotation", this, [this]() { emit textRotationChanged(0); });
    rotateBtn->setMenu(rotateMenu);
    bar->addWidget(rotateBtn);

    bar->addSeparator();

    // ===== Borders (Excel-style split button) =====
    QToolButton* borderBtn = new QToolButton(bar);
    borderBtn->setIcon(createBorderIcon());
    borderBtn->setToolTip("Borders");
    borderBtn->setPopupMode(QToolButton::MenuButtonPopup);
    borderBtn->setFixedSize(52, 24);
    borderBtn->setStyleSheet(
        "QToolButton { background: transparent; border: 1px solid transparent; border-radius: 4px; padding: 2px 6px; }"
        "QToolButton:hover { background-color: #E8ECF0; border-color: #D0D5DD; }"
        "QToolButton::menu-button { width: 14px; border-left: 1px solid #D0D5DD; }"
        "QToolButton::menu-button:hover { background-color: #D8DCE0; }"
    );

    QMenu* borderMenu = new QMenu(borderBtn);
    borderMenu->setStyleSheet(
        "QMenu { background: #FFFFFF; border: 1px solid #D0D5DD; border-radius: 6px; padding: 4px; }"
        "QMenu::item { padding: 5px 16px 5px 8px; border-radius: 4px; font-size: 12px; }"
        "QMenu::item:selected { background-color: #E8F0FE; }"
        "QMenu::icon { margin-right: 8px; }"
        "QMenu::separator { height: 1px; background: #E0E3E8; margin: 3px 8px; }"
    );

    auto emitBorder = [this](const QString& type) {
        emit borderStyleSelected(type, m_lastBorderColor, m_lastBorderWidth, m_lastBorderPenStyle);
    };

    borderMenu->addAction(createBorderMenuIcon("bottom"), "Bottom Border", this, [emitBorder]() { emitBorder("bottom"); });
    borderMenu->addAction(createBorderMenuIcon("top"), "Top Border", this, [emitBorder]() { emitBorder("top"); });
    borderMenu->addAction(createBorderMenuIcon("left"), "Left Border", this, [emitBorder]() { emitBorder("left"); });
    borderMenu->addAction(createBorderMenuIcon("right"), "Right Border", this, [emitBorder]() { emitBorder("right"); });
    borderMenu->addSeparator();
    borderMenu->addAction(createBorderMenuIcon("all"), "All Borders", this, [emitBorder]() { emitBorder("all"); });
    borderMenu->addAction(createBorderMenuIcon("outside"), "Outside Borders", this, [emitBorder]() { emitBorder("outside"); });
    borderMenu->addAction(createBorderMenuIcon("thick_outside"), "Thick Box Border", this, [emitBorder]() { emitBorder("thick_outside"); });
    borderMenu->addSeparator();
    borderMenu->addAction(createBorderMenuIcon("inside_h"), "Inside Horizontal", this, [emitBorder]() { emitBorder("inside_h"); });
    borderMenu->addAction(createBorderMenuIcon("inside_v"), "Inside Vertical", this, [emitBorder]() { emitBorder("inside_v"); });
    borderMenu->addAction(createBorderMenuIcon("inside"), "Inside Borders", this, [emitBorder]() { emitBorder("inside"); });
    borderMenu->addSeparator();
    borderMenu->addAction(createBorderMenuIcon("none"), "No Border", this, [emitBorder]() { emitBorder("none"); });

    borderMenu->addSeparator();

    // Helper: re-open border menu after Line Style / Line Color changes
    auto reopenBorderMenu = [borderMenu, borderBtn]() {
        QTimer::singleShot(0, borderMenu, [borderMenu, borderBtn]() {
            borderMenu->popup(borderBtn->mapToGlobal(QPoint(0, borderBtn->height())));
        });
    };

    // ===== Line Style submenu =====
    auto makeLineStyleIcon = [](int width, int style) {
        return createIcon(16, [width, style](QPainter& p, int) {
            QPen pen(QColor("#333"), width * 1.2);
            if (style == 1) pen.setStyle(Qt::DashLine);
            else if (style == 2) pen.setStyle(Qt::DotLine);
            p.setPen(pen);
            p.drawLine(1, 8, 15, 8);
        });
    };

    QMenu* lineStyleMenu = borderMenu->addMenu(
        makeLineStyleIcon(m_lastBorderWidth, m_lastBorderPenStyle),
        "Line Style: Thin");
    lineStyleMenu->setStyleSheet(borderMenu->styleSheet());

    // Collect all line style actions for exclusive checking
    QList<QAction*> lineActions;

    auto* thinAction = lineStyleMenu->addAction(makeLineStyleIcon(1, 0), "Thin");
    lineActions << thinAction;
    auto* mediumAction = lineStyleMenu->addAction(makeLineStyleIcon(2, 0), "Medium");
    lineActions << mediumAction;
    auto* thickAction = lineStyleMenu->addAction(makeLineStyleIcon(3, 0), "Thick");
    lineActions << thickAction;

    lineStyleMenu->addSeparator();

    auto* thinDashedAction = lineStyleMenu->addAction(makeLineStyleIcon(1, 1), "Thin Dashed");
    lineActions << thinDashedAction;
    auto* mediumDashedAction = lineStyleMenu->addAction(makeLineStyleIcon(2, 1), "Medium Dashed");
    lineActions << mediumDashedAction;

    lineStyleMenu->addSeparator();

    auto* thinDottedAction = lineStyleMenu->addAction(makeLineStyleIcon(1, 2), "Dotted");
    lineActions << thinDottedAction;

    for (auto* a : lineActions) a->setCheckable(true);
    thinAction->setChecked(true);

    auto uncheckAllLineStyles = [lineActions](QAction* except) {
        for (auto* a : lineActions) a->setChecked(a == except);
    };

    auto updateLineStyleLabel = [lineStyleMenu, makeLineStyleIcon](int width, int penStyle, const QString& name) {
        lineStyleMenu->setTitle(QString("Line Style: %1").arg(name));
        lineStyleMenu->setIcon(makeLineStyleIcon(width, penStyle));
    };

    connect(thinAction, &QAction::triggered, this, [this, uncheckAllLineStyles, thinAction, updateLineStyleLabel, reopenBorderMenu]() {
        m_lastBorderWidth = 1; m_lastBorderPenStyle = 0;
        uncheckAllLineStyles(thinAction);
        updateLineStyleLabel(1, 0, "Thin");
        reopenBorderMenu();
    });
    connect(mediumAction, &QAction::triggered, this, [this, uncheckAllLineStyles, mediumAction, updateLineStyleLabel, reopenBorderMenu]() {
        m_lastBorderWidth = 2; m_lastBorderPenStyle = 0;
        uncheckAllLineStyles(mediumAction);
        updateLineStyleLabel(2, 0, "Medium");
        reopenBorderMenu();
    });
    connect(thickAction, &QAction::triggered, this, [this, uncheckAllLineStyles, thickAction, updateLineStyleLabel, reopenBorderMenu]() {
        m_lastBorderWidth = 3; m_lastBorderPenStyle = 0;
        uncheckAllLineStyles(thickAction);
        updateLineStyleLabel(3, 0, "Thick");
        reopenBorderMenu();
    });
    connect(thinDashedAction, &QAction::triggered, this, [this, uncheckAllLineStyles, thinDashedAction, updateLineStyleLabel, reopenBorderMenu]() {
        m_lastBorderWidth = 1; m_lastBorderPenStyle = 1;
        uncheckAllLineStyles(thinDashedAction);
        updateLineStyleLabel(1, 1, "Thin Dashed");
        reopenBorderMenu();
    });
    connect(mediumDashedAction, &QAction::triggered, this, [this, uncheckAllLineStyles, mediumDashedAction, updateLineStyleLabel, reopenBorderMenu]() {
        m_lastBorderWidth = 2; m_lastBorderPenStyle = 1;
        uncheckAllLineStyles(mediumDashedAction);
        updateLineStyleLabel(2, 1, "Medium Dashed");
        reopenBorderMenu();
    });
    connect(thinDottedAction, &QAction::triggered, this, [this, uncheckAllLineStyles, thinDottedAction, updateLineStyleLabel, reopenBorderMenu]() {
        m_lastBorderWidth = 1; m_lastBorderPenStyle = 2;
        uncheckAllLineStyles(thinDottedAction);
        updateLineStyleLabel(1, 2, "Dotted");
        reopenBorderMenu();
    });

    // ===== Line Color =====
    auto makeColorSwatchIcon = [](const QColor& color) {
        return createIcon(16, [color](QPainter& p, int) {
            p.setRenderHint(QPainter::Antialiasing, true);
            p.setPen(QPen(QColor("#888"), 0.8));
            p.setBrush(color);
            p.drawRoundedRect(1, 1, 14, 14, 2, 2);
        });
    };
    auto* lineColorAction = borderMenu->addAction(
        makeColorSwatchIcon(m_lastBorderColor),
        QString("Line Color: %1").arg(m_lastBorderColor.name().toUpper()),
        this, [this, makeColorSwatchIcon, borderMenu, reopenBorderMenu]() {
            const DocumentTheme& dt = m_docTheme ? *m_docTheme : defaultDocumentTheme();
            auto pick = showColorPalette(this, m_lastBorderColorStr, dt, "Border Color");
            if (pick.isValid) {
                m_lastBorderColor = pick.displayColor;
                m_lastBorderColorStr = pick.colorString;
                // Update the action text and icon to show new color
                auto actions = borderMenu->actions();
                for (auto* a : actions) {
                    if (a->text().startsWith("Line Color:")) {
                        a->setText(QString("Line Color: %1").arg(pick.displayColor.name().toUpper()));
                        a->setIcon(makeColorSwatchIcon(pick.displayColor));
                        break;
                    }
                }
            }
            reopenBorderMenu();
        });

    borderBtn->setMenu(borderMenu);
    connect(borderBtn, &QToolButton::clicked, this, [emitBorder]() { emitBorder("all"); });
    bar->addWidget(borderBtn);

    // ===== Merge (Excel-style split button with text) =====
    QToolButton* mergeBtn = new QToolButton(bar);
    mergeBtn->setIcon(createMergeIcon());
    mergeBtn->setText("Merge");
    mergeBtn->setToolButtonStyle(Qt::ToolButtonTextBesideIcon);
    mergeBtn->setToolTip("Merge & Center");
    mergeBtn->setPopupMode(QToolButton::MenuButtonPopup);
    mergeBtn->setFixedHeight(24);
    mergeBtn->setStyleSheet(
        "QToolButton { background: transparent; border: 1px solid transparent; border-radius: 4px; "
        "padding: 2px 5px; font-size: 11px; color: #344054; }"
        "QToolButton:hover { background-color: #E8ECF0; border-color: #D0D5DD; }"
        "QToolButton::menu-button { width: 13px; border-left: 1px solid #D0D5DD; }"
        "QToolButton::menu-button:hover { background-color: #D8DCE0; }"
    );

    QMenu* mergeMenu = new QMenu(mergeBtn);
    mergeMenu->setStyleSheet(
        "QMenu { background: #FFFFFF; border: 1px solid #D0D5DD; border-radius: 6px; padding: 4px; }"
        "QMenu::item { padding: 5px 16px 5px 8px; border-radius: 4px; }"
        "QMenu::item:selected { background-color: #E8F0FE; }"
    );
    mergeMenu->addAction(createMergeIcon(), "Merge && Center", this, &Toolbar::mergeCellsRequested);
    mergeMenu->addAction("Unmerge Cells", this, &Toolbar::unmergeCellsRequested);

    mergeBtn->setMenu(mergeMenu);
    connect(mergeBtn, &QToolButton::clicked, this, &Toolbar::mergeCellsRequested);
    bar->addWidget(mergeBtn);

    bar->addSeparator();

    // ===== Number Format Dropdown (Google Sheets style) =====
    m_numberFormatBtn = new QToolButton(bar);
    m_numberFormatBtn->setText("General");
    m_numberFormatBtn->setToolTip("Number Format");
    m_numberFormatBtn->setPopupMode(QToolButton::InstantPopup);
    m_numberFormatBtn->setFixedHeight(24);
    m_numberFormatBtn->setMinimumWidth(80);
    m_numberFormatBtn->setToolButtonStyle(Qt::ToolButtonTextOnly);
    m_numberFormatBtn->setStyleSheet(
        "QToolButton { background: white; border: 1px solid #D0D5DD; border-radius: 4px; "
        "padding: 2px 18px 2px 8px; font-size: 11px; color: #344054; text-align: left; }"
        "QToolButton:hover { border-color: #4A90D9; }"
        "QToolButton::menu-indicator { subcontrol-position: right center; "
        "subcontrol-origin: padding; width: 10px; height: 10px; right: 4px; }"
    );

    // Build the number format menu
    QMenu* numFmtMenu = new QMenu(m_numberFormatBtn);
    numFmtMenu->setStyleSheet(
        "QMenu { background: #FFFFFF; border: 1px solid #D0D5DD; border-radius: 6px; padding: 4px; min-width: 220px; }"
        "QMenu::item { padding: 6px 16px 6px 12px; border-radius: 4px; font-size: 12px; }"
        "QMenu::item:selected { background-color: #E8F0FE; }"
        "QMenu::separator { height: 1px; background: #E0E3E8; margin: 3px 8px; }"
        "QMenu::item:disabled { color: #999; }"
    );

    auto addFmtItem = [&](const QString& label, const QString& shortcut, const QString& format) {
        QAction* action = numFmtMenu->addAction(label);
        if (!shortcut.isEmpty()) action->setShortcut(QKeySequence(shortcut));
        connect(action, &QAction::triggered, this, [this, format, label]() {
            emit numberFormatChanged(format);
            m_numberFormatBtn->setText(label);
        });
        return action;
    };

    addFmtItem("General", "", "General");
    addFmtItem("Number", "Ctrl+Shift+1", "Number");

    // --- Accounting submenu ---
    QMenu* accountingMenu = numFmtMenu->addMenu("Accounting");
    accountingMenu->setStyleSheet(numFmtMenu->styleSheet());
    for (const auto& cur : NumberFormat::currencies()) {
        accountingMenu->addAction(
            QString("%1  %2").arg(cur.symbol, -4).arg(cur.label),
            this, [this, cur]() {
                emit accountingFormatSelected(cur.code);
                m_numberFormatBtn->setText("Accounting");
            });
    }

    // --- Currency submenu ---
    QMenu* currencyMenu = numFmtMenu->addMenu("Currency");
    currencyMenu->setStyleSheet(numFmtMenu->styleSheet());
    for (const auto& cur : NumberFormat::currencies()) {
        currencyMenu->addAction(
            QString("%1  %2").arg(cur.symbol, -4).arg(cur.label),
            this, [this, cur]() {
                emit currencyFormatSelected(cur.code);
                m_numberFormatBtn->setText("Currency");
            });
    }

    // --- Date submenu with live previews ---
    QMenu* dateMenu = numFmtMenu->addMenu("Date");
    dateMenu->setStyleSheet(
        "QMenu { background: #FFFFFF; border: 1px solid #D0D5DD; border-radius: 6px; padding: 4px; min-width: 360px; }"
        "QMenu::item { padding: 6px 16px 6px 12px; border-radius: 4px; font-size: 12px; }"
        "QMenu::item:selected { background-color: #E8F0FE; }"
        "QMenu::separator { height: 1px; background: #E0E3E8; margin: 3px 8px; }"
        "QMenu::item:disabled { color: #888; font-size: 11px; font-weight: bold; }"
    );

    QDate today = QDate::currentDate();
    QDateTime now = QDateTime::currentDateTime();

    // Date section header
    auto* dateHeader = dateMenu->addAction("Date");
    dateHeader->setEnabled(false);
    dateMenu->addSeparator();

    struct DateFmt { QString preview; QString qtFmt; QString fmtId; QString label; };
    QList<DateFmt> dateFmts = {
        { today.toString("d/M/yy"),               "d/M/yy",                "d/M/yy",              "d/M/yy" },
        { today.toString("d MMM, yyyy"),           "d MMM, yyyy",           "d MMM, yyyy",         "d MMM, yyyy" },
        { today.toString("d MMMM, yyyy"),          "d MMMM, yyyy",         "d MMMM, yyyy",        "d MMMM, yyyy" },
        { today.toString("dddd, d MMMM, yyyy"),    "dddd, d MMMM, yyyy",   "EEEE, d MMMM, yyyy", "EEEE, d MMMM, yyyy" },
        { today.toString("dd/MM/yyyy"),            "dd/MM/yyyy",            "dd/MM/yyyy",          "dd/MM/yyyy" },
        { today.toString("MM/dd/yyyy"),            "MM/dd/yyyy",            "MM/dd/yyyy",          "MM/dd/yyyy" },
        { today.toString("yyyy/MM/dd"),            "yyyy/MM/dd",            "yyyy/MM/dd",          "yyyy/MM/dd" },
    };

    for (const auto& df : dateFmts) {
        QAction* action = dateMenu->addAction(QString("%1").arg(df.preview));
        // Show format code on the right via tooltip
        action->setToolTip(df.label);
        connect(action, &QAction::triggered, this, [this, df]() {
            emit dateFormatSelected(df.fmtId);
            m_numberFormatBtn->setText("Date");
        });
    }

    // Date and Time section
    dateMenu->addSeparator();
    auto* dtHeader = dateMenu->addAction("Date and Time");
    dtHeader->setEnabled(false);
    dateMenu->addSeparator();

    struct DateTimeFmt { QString fmtId; QString qtFmt; };
    QList<DateTimeFmt> dtFmts = {
        { "d/M/yy h:mm:ss a z",      "d/M/yy h:mm:ss AP t" },
        { "d MMM, yyyy h:mm:ss a z",  "d MMM, yyyy h:mm:ss AP t" },
        { "d MMMM, yyyy h:mm:ss a",   "d MMMM, yyyy h:mm:ss AP" },
        { "EEEE, d MMMM, yyyy h:mm a","dddd, d MMMM, yyyy h:mm AP" },
        { "d/M/yy h:mm a",            "d/M/yy h:mm AP" },
    };

    for (const auto& dtf : dtFmts) {
        QString preview = now.toString(dtf.qtFmt);
        QAction* action = dateMenu->addAction(preview);
        connect(action, &QAction::triggered, this, [this, dtf]() {
            emit dateFormatSelected(dtf.fmtId);
            m_numberFormatBtn->setText("Date");
        });
    }

    // --- Time ---
    addFmtItem("Time", "Ctrl+Shift+2", "Time");

    // --- Percentage ---
    addFmtItem("Percentage", "Ctrl+Shift+5", "Percentage");

    // --- Fraction ---
    addFmtItem("Fraction", "", "Fraction");

    // --- Scientific ---
    addFmtItem("Scientific", "Ctrl+Shift+6", "Scientific");

    // --- Text ---
    addFmtItem("Text", "", "Text");

    numFmtMenu->addSeparator();

    // --- More Formats... (opens Format Cells dialog) ---
    numFmtMenu->addAction("More Formats...", this, &Toolbar::formatCellsRequested);

    m_numberFormatBtn->setMenu(numFmtMenu);
    bar->addWidget(m_numberFormatBtn);

    // Quick-access currency button
    QToolButton* currencyBtn = new QToolButton(bar);
    currencyBtn->setText("\u20B9");  // ₹ (Indian Rupee — matches screenshot)
    currencyBtn->setFont(QFont("Arial", 12, QFont::Bold));
    currencyBtn->setToolTip("Currency Format");
    currencyBtn->setPopupMode(QToolButton::MenuButtonPopup);
    currencyBtn->setFixedSize(40, 24);
    currencyBtn->setStyleSheet(
        "QToolButton { background: transparent; border: 1px solid transparent; border-radius: 4px; padding: 2px 4px; }"
        "QToolButton:hover { background-color: #E8ECF0; border-color: #D0D5DD; }"
        "QToolButton::menu-button { width: 12px; border-left: 1px solid #D0D5DD; }"
        "QToolButton::menu-button:hover { background-color: #D8DCE0; }"
    );
    QMenu* quickCurrMenu = new QMenu(currencyBtn);
    quickCurrMenu->setStyleSheet(numFmtMenu->styleSheet());
    for (const auto& cur : NumberFormat::currencies()) {
        quickCurrMenu->addAction(
            QString("%1  %2").arg(cur.symbol, -4).arg(cur.label),
            this, [this, cur]() {
                emit currencyFormatSelected(cur.code);
                m_numberFormatBtn->setText("Currency");
            });
    }
    currencyBtn->setMenu(quickCurrMenu);
    bar->addWidget(currencyBtn);
    connect(currencyBtn, &QToolButton::clicked, this, [this]() {
        emit numberFormatChanged("Currency");
        m_numberFormatBtn->setText("Currency");
    });

    // Thousand separator (,)
    QToolButton* thousandBtn = new QToolButton(bar);
    thousandBtn->setText("(,)");
    thousandBtn->setFont(QFont("Arial", 10, QFont::Bold));
    thousandBtn->setToolTip("Thousand Separator");
    thousandBtn->setCheckable(true);
    thousandBtn->setFixedSize(28, 24);
    bar->addWidget(thousandBtn);
    connect(thousandBtn, &QToolButton::clicked, this, [this]() { emit thousandSeparatorToggled(); });

    // Increase decimals .00 →
    QToolButton* incDecBtn = new QToolButton(bar);
    incDecBtn->setToolTip("Increase Decimal Places");
    incDecBtn->setFixedSize(28, 24);
    incDecBtn->setIcon(createIcon(16, [](QPainter& p, int) {
        p.setRenderHint(QPainter::Antialiasing, true);
        QFont f("Arial", 6);
        p.setFont(f);
        p.setPen(QColor("#555"));
        p.drawText(QRect(0, 1, 12, 14), Qt::AlignCenter, ".00");
        // Arrow right
        p.setPen(QPen(QColor("#4A90D9"), 1.4, Qt::SolidLine, Qt::RoundCap));
        p.drawLine(11, 8, 15, 8);
        p.setPen(Qt::NoPen);
        p.setBrush(QColor("#4A90D9"));
        QPolygonF arr;
        arr << QPointF(14, 6) << QPointF(14, 10) << QPointF(16, 8);
        p.drawPolygon(arr);
    }));
    bar->addWidget(incDecBtn);
    connect(incDecBtn, &QToolButton::clicked, this, &Toolbar::increaseDecimals);

    // Decrease decimals .00 ←
    QToolButton* decDecBtn = new QToolButton(bar);
    decDecBtn->setToolTip("Decrease Decimal Places");
    decDecBtn->setFixedSize(28, 24);
    decDecBtn->setIcon(createIcon(16, [](QPainter& p, int) {
        p.setRenderHint(QPainter::Antialiasing, true);
        QFont f("Arial", 6);
        p.setFont(f);
        p.setPen(QColor("#555"));
        p.drawText(QRect(4, 1, 12, 14), Qt::AlignCenter, ".0");
        // Arrow left
        p.setPen(QPen(QColor("#4A90D9"), 1.4, Qt::SolidLine, Qt::RoundCap));
        p.drawLine(5, 8, 1, 8);
        p.setPen(Qt::NoPen);
        p.setBrush(QColor("#4A90D9"));
        QPolygonF arr;
        arr << QPointF(2, 6) << QPointF(2, 10) << QPointF(0, 8);
        p.drawPolygon(arr);
    }));
    bar->addWidget(decDecBtn);
    connect(decDecBtn, &QToolButton::clicked, this, &Toolbar::decreaseDecimals);

    bar->addSeparator();

    // ===== Data =====
    QToolButton* sortAscBtn = new QToolButton(bar);
    sortAscBtn->setIcon(createSortIcon(true));
    sortAscBtn->setToolTip("Sort A to Z");
    sortAscBtn->setFixedSize(28, 24);
    bar->addWidget(sortAscBtn);
    connect(sortAscBtn, &QToolButton::clicked, this, &Toolbar::sortAscending);

    QToolButton* sortDescBtn = new QToolButton(bar);
    sortDescBtn->setIcon(createSortIcon(false));
    sortDescBtn->setToolTip("Sort Z to A");
    sortDescBtn->setFixedSize(28, 24);
    bar->addWidget(sortDescBtn);
    connect(sortDescBtn, &QToolButton::clicked, this, &Toolbar::sortDescending);

    QToolButton* filterBtn = new QToolButton(bar);
    filterBtn->setIcon(createFilterIcon());
    filterBtn->setToolTip("Toggle Auto Filter");
    filterBtn->setCheckable(true);
    filterBtn->setFixedSize(28, 24);
    bar->addWidget(filterBtn);
    connect(filterBtn, &QToolButton::clicked, this, &Toolbar::filterToggled);

    bar->addSeparator();

    // ===== Table =====
    QToolButton* tableBtn = new QToolButton(bar);
    tableBtn->setIcon(createTableIcon());
    tableBtn->setToolTip("Format as Table");
    tableBtn->setPopupMode(QToolButton::InstantPopup);
    tableBtn->setFixedSize(38, 24);
    tableBtn->setStyleSheet(
        "QToolButton { background: transparent; border: 1px solid transparent; border-radius: 4px; padding: 2px 4px; }"
        "QToolButton:hover { background-color: #E8ECF0; border-color: #D0D5DD; }"
        "QToolButton::menu-indicator { subcontrol-position: right center; width: 8px; height: 8px; }"
    );

    QMenu* tableMenu = new QMenu(tableBtn);
    tableMenu->setStyleSheet(
        "QMenu { background: #FFFFFF; border: 1px solid #D0D5DD; border-radius: 6px; padding: 4px; }"
        "QMenu::item { padding: 6px 14px 6px 10px; border-radius: 4px; }"
        "QMenu::item:selected { background-color: #E8F0FE; }"
    );

    // Generate table themes from document theme (re-generated each time menu opens)
    connect(tableMenu, &QMenu::aboutToShow, this, [this, tableMenu]() {
        tableMenu->clear();
        const DocumentTheme& dt = m_docTheme ? *m_docTheme : defaultDocumentTheme();
        auto themes = generateTableThemes(dt);
        for (int i = 0; i < static_cast<int>(themes.size()); ++i) {
            const auto& theme = themes[i];
            QPixmap swatch(48, 28);
            swatch.fill(Qt::transparent);
            QPainter sp(&swatch);
            sp.setRenderHint(QPainter::Antialiasing, true);
            QPainterPath clip;
            clip.addRoundedRect(0, 0, 48, 28, 4, 4);
            sp.setClipPath(clip);
            sp.fillRect(0, 0, 48, 8, theme.headerBg);
            sp.fillRect(0, 8, 48, 7, theme.bandedRow1);
            sp.fillRect(0, 15, 48, 6, theme.bandedRow2);
            sp.fillRect(0, 21, 48, 7, theme.bandedRow1);
            sp.setClipping(false);
            sp.setPen(QPen(QColor("#D0D5DD"), 0.5));
            sp.drawRoundedRect(QRectF(0.25, 0.25, 47.5, 27.5), 4, 4);
            sp.end();
            QAction* action = tableMenu->addAction(QIcon(swatch), theme.name);
            connect(action, &QAction::triggered, this, [this, i]() { emit tableStyleSelected(i); });
        }
    });
    tableBtn->setMenu(tableMenu);
    bar->addWidget(tableBtn);

    bar->addSeparator();

    // ===== Conditional Formatting =====
    QToolButton* condFmtBtn = new QToolButton(bar);
    condFmtBtn->setIcon(createCondFmtIcon());
    condFmtBtn->setToolTip("Conditional Formatting");
    condFmtBtn->setFixedSize(28, 24);
    bar->addWidget(condFmtBtn);
    connect(condFmtBtn, &QToolButton::clicked, this, &Toolbar::conditionalFormatRequested);

    // ===== Data Validation =====
    QToolButton* validationBtn = new QToolButton(bar);
    validationBtn->setIcon(createValidationIcon());
    validationBtn->setToolTip("Data Validation");
    validationBtn->setFixedSize(28, 24);
    bar->addWidget(validationBtn);
    connect(validationBtn, &QToolButton::clicked, this, &Toolbar::dataValidationRequested);

    bar->addSeparator();

    // ===== Checkbox =====
    QToolButton* checkboxBtn = new QToolButton(bar);
    checkboxBtn->setIcon(createIcon(16, [](QPainter& p, int) {
        p.setRenderHint(QPainter::Antialiasing, true);
        const QColor accent = ThemeManager::instance().currentTheme().accentDark;
        p.setPen(QPen(accent, 1.5));
        p.setBrush(Qt::white);
        p.drawRoundedRect(2, 2, 12, 12, 2, 2);
        p.setPen(QPen(accent, 2.0, Qt::SolidLine, Qt::RoundCap, Qt::RoundJoin));
        p.drawLine(4, 8, 6, 11);
        p.drawLine(6, 11, 12, 4);
    }));
    checkboxBtn->setToolButtonStyle(Qt::ToolButtonIconOnly);
    checkboxBtn->setToolTip("Insert Checkbox");
    checkboxBtn->setFixedHeight(24);
    bar->addWidget(checkboxBtn);
    connect(checkboxBtn, &QToolButton::clicked, this, &Toolbar::insertCheckboxRequested);

    // ===== Picklist (dropdown menu: Insert + Manage) =====
    QToolButton* picklistBtn = new QToolButton(bar);
    picklistBtn->setIcon(createIcon(16, [](QPainter& p, int) {
        p.setRenderHint(QPainter::Antialiasing, true);
        p.setPen(Qt::NoPen);
        p.setBrush(QColor("#DBEAFE"));
        p.drawRoundedRect(1, 1, 14, 6, 3, 3);
        p.setPen(QColor("#1E40AF"));
        QFont f = p.font(); f.setPixelSize(5); p.setFont(f);
        p.drawText(QRect(1, 1, 14, 6), Qt::AlignCenter, "Tag");
        p.setPen(Qt::NoPen);
        p.setBrush(QColor("#FCE7F3"));
        p.drawRoundedRect(1, 9, 14, 6, 3, 3);
        p.setPen(QColor("#9D174D"));
        p.drawText(QRect(1, 9, 14, 6), Qt::AlignCenter, "Tag");
    }));
    picklistBtn->setToolButtonStyle(Qt::ToolButtonTextBesideIcon);
    picklistBtn->setText("Picklist");
    picklistBtn->setToolTip("Picklist");
    picklistBtn->setFixedHeight(24);
    picklistBtn->setPopupMode(QToolButton::MenuButtonPopup);
    QMenu* picklistMenu = new QMenu(picklistBtn);
    picklistMenu->addAction("Insert Picklist", this, &Toolbar::insertPicklistRequested);
    picklistMenu->addAction("Manage Picklists...", this, &Toolbar::managePicklistsRequested);
    picklistBtn->setMenu(picklistMenu);
    connect(picklistBtn, &QToolButton::clicked, this, &Toolbar::insertPicklistRequested);
    bar->addWidget(picklistBtn);

    bar->addSeparator();

    // ===== Insert Chart =====
    QToolButton* chartBtn = new QToolButton(bar);
    chartBtn->setIcon(createIcon(16, [](QPainter& p, int) {
        // Column chart icon
        QColor c("#4A90D9");
        p.setPen(Qt::NoPen);
        p.setBrush(c);
        p.drawRect(2, 9, 3, 5);
        p.drawRect(6, 5, 3, 9);
        p.drawRect(10, 7, 3, 7);
        // Axis lines
        p.setPen(QPen(QColor("#555"), 1));
        p.drawLine(1, 14, 14, 14);
        p.drawLine(1, 2, 1, 14);
    }));
    chartBtn->setText("Chart");
    chartBtn->setToolButtonStyle(Qt::ToolButtonTextBesideIcon);
    chartBtn->setToolTip("Insert Chart (Alt+F1)");
    chartBtn->setFixedHeight(24);
    bar->addWidget(chartBtn);
    connect(chartBtn, &QToolButton::clicked, this, &Toolbar::insertChartRequested);

    // ===== Insert Shape =====
    QToolButton* shapeBtn = new QToolButton(bar);
    shapeBtn->setIcon(createIcon(16, [](QPainter& p, int) {
        // Shapes icon: circle + rectangle
        p.setPen(QPen(QColor("#4A90D9"), 1.4));
        p.setBrush(QColor("#4A90D9").lighter(170));
        p.drawRoundedRect(1, 5, 9, 9, 2, 2);
        p.drawEllipse(7, 1, 8, 8);
    }));
    shapeBtn->setText("Shape");
    shapeBtn->setToolButtonStyle(Qt::ToolButtonTextBesideIcon);
    shapeBtn->setToolTip("Insert Shape");
    shapeBtn->setFixedHeight(24);
    bar->addWidget(shapeBtn);
    connect(shapeBtn, &QToolButton::clicked, this, &Toolbar::insertShapeRequested);

    bar->addSeparator();

    // ===== Chat Assistant (at end of row 2) =====
    QToolButton* chatBtn = new QToolButton(bar);
    chatBtn->setIcon(createChatIcon());
    chatBtn->setText("Claude");
    chatBtn->setToolButtonStyle(Qt::ToolButtonTextBesideIcon);
    chatBtn->setToolTip("Open Claude Assistant");
    chatBtn->setFixedHeight(24);
    {
        const auto& t = ThemeManager::instance().currentTheme();
        chatBtn->setStyleSheet(QString(
            "QToolButton { background: %1; color: white; border: none; border-radius: 4px; "
            "padding: 2px 10px; font-size: 11px; font-weight: bold; }"
            "QToolButton:hover { background: %2; }"
            "QToolButton:pressed { background: %3; }"
        ).arg(t.accentDarker.name(), t.accentDark.name(), t.menuBarHover.name()));
    }
    bar->addWidget(chatBtn);
    connect(chatBtn, &QToolButton::clicked, this, &Toolbar::chatToggleRequested);

    m_secondaryToolbar = bar;
    return bar;
}

void Toolbar::onThemeChanged() {
    setStyleSheet(buildToolbarStyle());
    if (m_secondaryToolbar) {
        m_secondaryToolbar->setStyleSheet(buildToolbarStyleRow2());
    }
}

void Toolbar::updateBgColorIcon() {
    if (!m_bgColorBtn) return;
    m_bgColorBtn->setIcon(createIcon(16, [this](QPainter& p, int) {
        p.setRenderHint(QPainter::Antialiasing, true);
        p.setPen(Qt::NoPen);

        // Paint drop — single clean teardrop in active color
        p.setBrush(m_lastBgColor);
        QPainterPath drop;
        drop.moveTo(8, 1);
        drop.cubicTo(3, 6, 2, 10, 5, 12);
        drop.cubicTo(6.5, 13, 9.5, 13, 11, 12);
        drop.cubicTo(14, 10, 13, 6, 8, 1);
        p.drawPath(drop);

        // Thin outline to keep it visible on white
        p.setPen(QPen(QColor("#666"), 0.6));
        p.setBrush(Qt::NoBrush);
        p.drawPath(drop);

        // Color bar at bottom
        p.setRenderHint(QPainter::Antialiasing, false);
        p.setPen(Qt::NoPen);
        p.setBrush(m_lastBgColor);
        p.drawRect(1, 14, 14, 2);
    }));
    m_bgColorBtn->setStyleSheet(
        "QToolButton { background: transparent; border: 1px solid transparent; border-radius: 4px; }"
        "QToolButton:hover { background-color: #E8ECF0; border-color: #D0D5DD; }");
}

void Toolbar::syncToStyle(const CellStyle& style) {
    // Block signals so we don't trigger formatting changes on the cell
    m_fontCombo->blockSignals(true);
    m_fontSizeSpinBox->blockSignals(true);

    // Font family and size
    m_fontCombo->setCurrentFont(QFont(style.fontName));
    m_fontSizeSpinBox->setValue(style.fontSize);

    m_fontCombo->blockSignals(false);
    m_fontSizeSpinBox->blockSignals(false);

    // Bold / Italic / Underline / Strikethrough
    if (m_boldBtn) m_boldBtn->setChecked(style.bold);
    if (m_italicBtn) m_italicBtn->setChecked(style.italic);
    if (m_underlineBtn) m_underlineBtn->setChecked(style.underline);
    if (m_strikethroughBtn) m_strikethroughBtn->setChecked(style.strikethrough);

    // Font color
    if (m_fgColorBtn) {
        m_lastFgColorStr = style.foregroundColor;
        const DocumentTheme& dt = m_docTheme ? *m_docTheme : defaultDocumentTheme();
        m_lastFgColor = dt.resolveAnyColor(style.foregroundColor);
        m_fgColorBtn->setStyleSheet(
            QString("QToolButton { color: %1; font-weight: bold; border-bottom: 3px solid %1; border-radius: 4px; }")
                .arg(m_lastFgColor.name()));
    }

    // Fill color
    if (m_bgColorBtn) {
        m_lastBgColorStr = style.backgroundColor;
        const DocumentTheme& dt = m_docTheme ? *m_docTheme : defaultDocumentTheme();
        m_lastBgColor = dt.resolveAnyColor(style.backgroundColor);
        updateBgColorIcon();
    }

    // Horizontal alignment
    if (m_alignLeftBtn) m_alignLeftBtn->setChecked(style.hAlign == HorizontalAlignment::Left);
    if (m_alignCenterBtn) m_alignCenterBtn->setChecked(style.hAlign == HorizontalAlignment::Center);
    if (m_alignRightBtn) m_alignRightBtn->setChecked(style.hAlign == HorizontalAlignment::Right);

    // Vertical alignment
    if (m_vAlignTopBtn) m_vAlignTopBtn->setChecked(style.vAlign == VerticalAlignment::Top);
    if (m_vAlignMiddleBtn) m_vAlignMiddleBtn->setChecked(style.vAlign == VerticalAlignment::Middle);
    if (m_vAlignBottomBtn) m_vAlignBottomBtn->setChecked(style.vAlign == VerticalAlignment::Bottom);

    // Number format dropdown text
    if (m_numberFormatBtn) {
        m_numberFormatBtn->setText(style.numberFormat == "General" ? "General" : style.numberFormat);
    }
}
