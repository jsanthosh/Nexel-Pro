#include "Toolbar.h"
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
#include <functional>
#include "../core/TableStyle.h"

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
        // Speech bubble
        p.setPen(QPen(QColor("#1B5E3B"), 1.2));
        p.setBrush(QColor("#1B5E3B").lighter(160));
        QPainterPath bubble;
        bubble.addRoundedRect(2, 2, 14, 10, 3, 3);
        // Tail
        bubble.moveTo(5, 12);
        bubble.lineTo(4, 16);
        bubble.lineTo(9, 12);
        p.drawPath(bubble);
        // Dots
        p.setPen(Qt::NoPen);
        p.setBrush(QColor("#1B5E3B"));
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

static const char* TOOLBAR_STYLE = R"(
    QToolBar {
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 #1B5E3B, stop:0.04 #1B5E3B,
            stop:0.041 #FAFBFC, stop:1 #F0F2F5);
        border-bottom: 1px solid #D0D5DD;
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
        color: #344054;
    }
    QToolButton:hover {
        background-color: #E8ECF0;
        border-color: #D0D5DD;
    }
    QToolButton:pressed {
        background-color: #D0D5DD;
    }
    QToolButton:checked {
        background-color: #D6E4F0;
        border-color: #4A90D9;
    }
    QFontComboBox {
        max-width: 180px;
        min-width: 140px;
        height: 26px;
        border: 1px solid #D0D5DD;
        border-radius: 4px;
        padding: 1px 4px 1px 6px;
        background: white;
        font-size: 12px;
        color: #344054;
    }
    QFontComboBox:focus {
        border-color: #4A90D9;
    }
    QFontComboBox QAbstractItemView {
        min-width: 200px;
    }
    QSpinBox {
        max-width: 50px;
        height: 26px;
        border: 1px solid #D0D5DD;
        border-radius: 4px;
        padding: 1px 6px;
        background: white;
        font-size: 12px;
        color: #344054;
    }
    QSpinBox:focus {
        border-color: #4A90D9;
    }
    QToolBar::separator {
        width: 1px;
        background-color: #E0E3E8;
        margin: 4px 4px;
    }
)";

Toolbar::Toolbar(QWidget* parent)
    : QToolBar("Standard Toolbar", parent) {

    setMovable(false);
    setFloatable(false);
    setIconSize(QSize(18, 18));
    setStyleSheet(TOOLBAR_STYLE);

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

    QToolButton* saveBtn = new QToolButton(this);
    saveBtn->setIcon(createSaveIcon());
    saveBtn->setToolTip("Save Document (Ctrl+S)");
    saveBtn->setFixedSize(30, 28);
    addWidget(saveBtn);
    connect(saveBtn, &QToolButton::clicked, this, &Toolbar::saveDocument);

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
    makeFmtBtn("B", boldFont, "Bold (Ctrl+B)", QKeySequence::Bold, [this]() { emit bold(); });

    QFont italicFont("Arial", 12);
    italicFont.setItalic(true);
    makeFmtBtn("I", italicFont, "Italic (Ctrl+I)", QKeySequence::Italic, [this]() { emit italic(); });

    QFont underlineFont("Arial", 12);
    underlineFont.setUnderline(true);
    makeFmtBtn("U", underlineFont, "Underline (Ctrl+U)", QKeySequence::Underline, [this]() { emit underline(); });

    QFont strikeFont("Arial", 12);
    strikeFont.setStrikeOut(true);
    makeFmtBtn("S", strikeFont, "Strikethrough", QKeySequence(), [this]() { emit strikethrough(); });

    addSeparator();

    // ===== Colors =====
    QToolButton* fgColorBtn = new QToolButton(this);
    fgColorBtn->setText("A");
    fgColorBtn->setFont(QFont("Arial", 12, QFont::Bold));
    fgColorBtn->setToolTip("Font Color");
    fgColorBtn->setFixedSize(28, 28);
    fgColorBtn->setStyleSheet(
        "QToolButton { color: #C00000; font-weight: bold; border-bottom: 3px solid #C00000; border-radius: 4px; }");
    addWidget(fgColorBtn);
    connect(fgColorBtn, &QToolButton::clicked, this, [this, fgColorBtn]() {
        QColor color = QColorDialog::getColor(m_lastFgColor, this, "Font Color");
        if (color.isValid()) {
            m_lastFgColor = color;
            fgColorBtn->setStyleSheet(
                QString("QToolButton { color: %1; font-weight: bold; border-bottom: 3px solid %1; border-radius: 4px; }")
                    .arg(color.name()));
            emit foregroundColorChanged(color);
        }
    });

    QToolButton* bgColorBtn = new QToolButton(this);
    bgColorBtn->setToolTip("Fill Color");
    bgColorBtn->setFixedSize(28, 28);
    bgColorBtn->setStyleSheet(
        "QToolButton { background-color: #FFFF00; border: 1px solid #D0D5DD; border-bottom: 3px solid #FFFF00; border-radius: 4px; }");
    addWidget(bgColorBtn);
    // Paint a bucket icon on it
    {
        QPixmap pix(18, 18);
        pix.fill(Qt::transparent);
        QPainter p(&pix);
        p.setRenderHint(QPainter::Antialiasing, true);
        p.setPen(QPen(QColor("#555"), 1.0));
        p.setBrush(QColor("#FFFF00"));
        p.drawRoundedRect(3, 6, 12, 9, 2, 2);
        p.setPen(QPen(QColor("#888"), 0.8));
        p.drawLine(5, 6, 5, 3);
        p.drawLine(5, 3, 12, 3);
        p.end();
        bgColorBtn->setIcon(QIcon(pix));
    }
    connect(bgColorBtn, &QToolButton::clicked, this, [this, bgColorBtn]() {
        QColor color = QColorDialog::getColor(m_lastBgColor, this, "Fill Color");
        if (color.isValid()) {
            m_lastBgColor = color;
            bgColorBtn->setStyleSheet(
                QString("QToolButton { background-color: %1; border: 1px solid #D0D5DD; border-bottom: 3px solid %1; border-radius: 4px; }")
                    .arg(color.name()));
            emit backgroundColorChanged(color);
        }
    });

}

// ===== Row 2: Secondary Toolbar =====

static const char* TOOLBAR_STYLE_ROW2 = R"(
    QToolBar {
        background: #F0F2F5;
        border-bottom: 1px solid #D0D5DD;
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
        color: #344054;
    }
    QToolButton:hover {
        background-color: #E8ECF0;
        border-color: #D0D5DD;
    }
    QToolButton:pressed {
        background-color: #D0D5DD;
    }
    QToolButton:checked {
        background-color: #D6E4F0;
        border-color: #4A90D9;
    }
    QToolButton::menu-indicator {
        width: 0px;
        height: 0px;
        image: none;
    }
    QToolBar::separator {
        width: 1px;
        background-color: #E0E3E8;
        margin: 3px 4px;
    }
)";

QToolBar* Toolbar::createSecondaryToolbar(QWidget* parent) {
    QToolBar* bar = new QToolBar("Format Toolbar", parent);
    bar->setMovable(false);
    bar->setFloatable(false);
    bar->setIconSize(QSize(16, 16));
    bar->setStyleSheet(TOOLBAR_STYLE_ROW2);

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
        "QToolButton::menu-arrow { image: none; }"
    );

    QMenu* borderMenu = new QMenu(borderBtn);
    borderMenu->setStyleSheet(
        "QMenu { background: #FFFFFF; border: 1px solid #D0D5DD; border-radius: 6px; padding: 4px; }"
        "QMenu::item { padding: 5px 16px 5px 8px; border-radius: 4px; font-size: 12px; }"
        "QMenu::item:selected { background-color: #E8F0FE; }"
        "QMenu::icon { margin-right: 8px; }"
        "QMenu::separator { height: 1px; background: #E0E3E8; margin: 3px 8px; }"
    );
    borderMenu->addAction(createBorderMenuIcon("bottom"), "Bottom Border", this, [this]() { emit borderStyleSelected("bottom"); });
    borderMenu->addAction(createBorderMenuIcon("top"), "Top Border", this, [this]() { emit borderStyleSelected("top"); });
    borderMenu->addAction(createBorderMenuIcon("left"), "Left Border", this, [this]() { emit borderStyleSelected("left"); });
    borderMenu->addAction(createBorderMenuIcon("right"), "Right Border", this, [this]() { emit borderStyleSelected("right"); });
    borderMenu->addSeparator();
    borderMenu->addAction(createBorderMenuIcon("all"), "All Borders", this, [this]() { emit borderStyleSelected("all"); });
    borderMenu->addAction(createBorderMenuIcon("outside"), "Outside Borders", this, [this]() { emit borderStyleSelected("outside"); });
    borderMenu->addAction(createBorderMenuIcon("thick_outside"), "Thick Box Border", this, [this]() { emit borderStyleSelected("thick_outside"); });
    borderMenu->addSeparator();
    borderMenu->addAction(createBorderMenuIcon("inside_h"), "Inside Horizontal", this, [this]() { emit borderStyleSelected("inside_h"); });
    borderMenu->addAction(createBorderMenuIcon("inside_v"), "Inside Vertical", this, [this]() { emit borderStyleSelected("inside_v"); });
    borderMenu->addAction(createBorderMenuIcon("inside"), "Inside Borders", this, [this]() { emit borderStyleSelected("inside"); });
    borderMenu->addSeparator();
    borderMenu->addAction(createBorderMenuIcon("none"), "No Border", this, [this]() { emit borderStyleSelected("none"); });

    borderBtn->setMenu(borderMenu);
    connect(borderBtn, &QToolButton::clicked, this, [this]() { emit borderStyleSelected("all"); });
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
        "QToolButton::menu-arrow { image: none; }"
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

    // ===== Number formatting =====
    QToolButton* currencyBtn = new QToolButton(bar);
    currencyBtn->setText("$");
    currencyBtn->setFont(QFont("Arial", 12, QFont::Bold));
    currencyBtn->setToolTip("Currency Format");
    currencyBtn->setFixedSize(26, 24);
    bar->addWidget(currencyBtn);
    connect(currencyBtn, &QToolButton::clicked, this, [this]() { emit numberFormatChanged("Currency"); });

    QToolButton* percentBtn = new QToolButton(bar);
    percentBtn->setText("%");
    percentBtn->setFont(QFont("Arial", 12, QFont::Bold));
    percentBtn->setToolTip("Percentage Format");
    percentBtn->setFixedSize(26, 24);
    bar->addWidget(percentBtn);
    connect(percentBtn, &QToolButton::clicked, this, [this]() { emit numberFormatChanged("Percentage"); });

    QToolButton* thousandBtn = new QToolButton(bar);
    thousandBtn->setText(",");
    thousandBtn->setFont(QFont("Arial", 13, QFont::Bold));
    thousandBtn->setToolTip("Thousand Separator");
    thousandBtn->setCheckable(true);
    thousandBtn->setFixedSize(26, 24);
    bar->addWidget(thousandBtn);
    connect(thousandBtn, &QToolButton::clicked, this, [this]() { emit thousandSeparatorToggled(); });

    QToolButton* formatCellsBtn = new QToolButton(bar);
    formatCellsBtn->setIcon(createFormatCellsIcon());
    formatCellsBtn->setToolTip("Format Cells (Ctrl+1)");
    formatCellsBtn->setFixedSize(28, 24);
    bar->addWidget(formatCellsBtn);
    connect(formatCellsBtn, &QToolButton::clicked, this, &Toolbar::formatCellsRequested);

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
    tableBtn->setFixedSize(28, 24);

    QMenu* tableMenu = new QMenu(tableBtn);
    tableMenu->setStyleSheet(
        "QMenu { background: #FFFFFF; border: 1px solid #D0D5DD; border-radius: 6px; padding: 4px; }"
        "QMenu::item { padding: 6px 14px 6px 10px; border-radius: 4px; }"
        "QMenu::item:selected { background-color: #E8F0FE; }"
    );

    auto themes = getBuiltinTableThemes();
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
    chatBtn->setStyleSheet(
        "QToolButton { background: #1B5E3B; color: white; border: none; border-radius: 4px; "
        "padding: 2px 10px; font-size: 11px; font-weight: bold; }"
        "QToolButton:hover { background: #246B45; }"
        "QToolButton:pressed { background: #155030; }"
    );
    bar->addWidget(chatBtn);
    connect(chatBtn, &QToolButton::clicked, this, &Toolbar::chatToggleRequested);

    return bar;
}
