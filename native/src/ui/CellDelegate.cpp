#include "CellDelegate.h"
#include <QLineEdit>
#include <QPainter>
#include <QPainterPath>
#include <QStyleOptionViewItem>
#include <QApplication>
#include <QCompleter>
#include <QStringListModel>
#include <QAbstractItemView>
#include <QListView>
#include <QKeyEvent>
#include <QTableView>
#include <QTimer>

// All supported formula function names
static const QStringList s_formulaNames = {
    "SUM", "AVERAGE", "COUNT", "COUNTA", "MIN", "MAX",
    "IF", "IFERROR", "AND", "OR", "NOT",
    "CONCAT", "CONCATENATE", "LEN", "UPPER", "LOWER", "TRIM",
    "LEFT", "RIGHT", "MID", "FIND", "SUBSTITUTE", "TEXT",
    "ROUND", "ABS", "SQRT", "POWER", "MOD", "INT", "CEILING", "FLOOR",
    "COUNTIF", "SUMIF", "AVERAGEIF", "COUNTBLANK", "SUMPRODUCT",
    "MEDIAN", "MODE", "STDEV", "VAR", "LARGE", "SMALL", "RANK", "PERCENTILE",
    "NOW", "TODAY", "YEAR", "MONTH", "DAY",
    "DATE", "HOUR", "MINUTE", "SECOND", "DATEDIF", "NETWORKDAYS", "WEEKDAY",
    "EDATE", "EOMONTH", "DATEVALUE",
    "VLOOKUP", "HLOOKUP", "XLOOKUP", "INDEX", "MATCH",
    "ROUNDUP", "ROUNDDOWN", "LOG", "LN", "EXP", "RAND", "RANDBETWEEN",
    "PROPER", "SEARCH", "REPT", "EXACT", "VALUE",
    "ISBLANK", "ISERROR", "ISNUMBER", "ISTEXT", "CHOOSE", "SWITCH",
};

CellDelegate::CellDelegate(QObject* parent)
    : QStyledItemDelegate(parent) {
}

bool CellDelegate::eventFilter(QObject* object, QEvent* event) {
    if (event->type() == QEvent::KeyPress) {
        QKeyEvent* keyEvent = static_cast<QKeyEvent*>(event);
        // Arrow keys during editing: commit and move (like Excel)
        if (keyEvent->key() == Qt::Key_Up || keyEvent->key() == Qt::Key_Down ||
            keyEvent->key() == Qt::Key_Left || keyEvent->key() == Qt::Key_Right) {
            QLineEdit* editor = qobject_cast<QLineEdit*>(object);
            if (editor) {
                // Left/Right: only commit if cursor is at boundary
                if (keyEvent->key() == Qt::Key_Left && editor->cursorPosition() > 0)
                    return QStyledItemDelegate::eventFilter(object, event);
                if (keyEvent->key() == Qt::Key_Right && editor->cursorPosition() < editor->text().length())
                    return QStyledItemDelegate::eventFilter(object, event);

                emit commitData(editor);
                emit closeEditor(editor, QAbstractItemDelegate::NoHint);

                // Navigate to adjacent cell after editor closes
                int key = keyEvent->key();
                QWidget* viewport = editor->parentWidget();
                QTableView* view = viewport ? qobject_cast<QTableView*>(viewport->parentWidget()) : nullptr;
                if (view) {
                    QTimer::singleShot(0, view, [view, key]() {
                        QModelIndex cur = view->currentIndex();
                        if (!cur.isValid() || !view->model()) return;
                        int row = cur.row(), col = cur.column();
                        if (key == Qt::Key_Up) row = qMax(0, row - 1);
                        else if (key == Qt::Key_Down) row = qMin(view->model()->rowCount() - 1, row + 1);
                        else if (key == Qt::Key_Left) col = qMax(0, col - 1);
                        else if (key == Qt::Key_Right) col = qMin(view->model()->columnCount() - 1, col + 1);
                        QModelIndex next = view->model()->index(row, col);
                        if (next.isValid()) {
                            view->setCurrentIndex(next);
                        }
                    });
                }
                return true; // consume the event
            }
        }
    }
    return QStyledItemDelegate::eventFilter(object, event);
}

QWidget* CellDelegate::createEditor(QWidget* parent, const QStyleOptionViewItem& option,
                                   const QModelIndex& index) const {
    QLineEdit* editor = new QLineEdit(parent);
    editor->setFrame(false);
    editor->setStyleSheet(
        "QLineEdit { background: white; padding: 1px 2px; "
        "border: 2px solid #107C10; selection-background-color: #0078D4; }");

    // Formula autocomplete
    auto* completer = new QCompleter(s_formulaNames, editor);
    completer->setWidget(editor);
    completer->setCaseSensitivity(Qt::CaseInsensitive);
    completer->setCompletionMode(QCompleter::PopupCompletion);
    completer->setFilterMode(Qt::MatchContains);
    completer->popup()->setStyleSheet(
        "QListView { background: white; border: 1px solid #C8C8C8; font-size: 12px; }"
        "QListView::item { padding: 3px 8px; }"
        "QListView::item:selected { background: #E8F0FE; color: #333; }"
    );

    // When a formula is selected from the completer, append parentheses
    connect(completer, QOverload<const QString&>::of(&QCompleter::activated),
            editor, [editor](const QString& funcName) {
        // The completer replaces text after '=', so we need to set it with parens
        QString text = editor->text();
        // Find where the function name starts (after '=' and any existing content)
        int eqPos = text.lastIndexOf('=');
        if (eqPos >= 0) {
            editor->setText(text.left(eqPos + 1) + funcName + "()");
            editor->setCursorPosition(editor->text().length() - 1); // Position between parens
        }
    });

    // Show completer only when typing after '='
    connect(editor, &QLineEdit::textChanged, this, [this, editor, completer](const QString& text) {
        emit formulaEditModeChanged(text.startsWith("="));

        if (text.startsWith("=") && text.length() > 1) {
            // Extract the token being typed (after last operator/paren/comma)
            QString afterEq = text.mid(1);
            int lastDelim = -1;
            for (int i = afterEq.length() - 1; i >= 0; --i) {
                QChar ch = afterEq[i];
                if (ch == '(' || ch == ')' || ch == ',' || ch == '+' || ch == '-' ||
                    ch == '*' || ch == '/' || ch == ':' || ch == ' ') {
                    lastDelim = i;
                    break;
                }
            }
            QString prefix = afterEq.mid(lastDelim + 1);
            if (!prefix.isEmpty() && prefix[0].isLetter()) {
                completer->setCompletionPrefix(prefix);
                if (completer->completionCount() > 0) {
                    completer->complete();
                } else {
                    completer->popup()->hide();
                }
            } else {
                completer->popup()->hide();
            }
        } else {
            completer->popup()->hide();
        }
    });

    return editor;
}

void CellDelegate::setEditorData(QWidget* editor, const QModelIndex& index) const {
    QLineEdit* lineEdit = qobject_cast<QLineEdit*>(editor);
    if (lineEdit) {
        lineEdit->setText(index.data(Qt::EditRole).toString());
        // Place cursor at end (not selecting all text) — like Excel
        lineEdit->deselect();
        lineEdit->setCursorPosition(lineEdit->text().length());
    }
}

void CellDelegate::setModelData(QWidget* editor, QAbstractItemModel* model,
                               const QModelIndex& index) const {
    QLineEdit* lineEdit = qobject_cast<QLineEdit*>(editor);
    if (lineEdit) {
        model->setData(index, lineEdit->text(), Qt::EditRole);
    }
}

void CellDelegate::updateEditorGeometry(QWidget* editor, const QStyleOptionViewItem& option,
                                       const QModelIndex& index) const {
    // Editor fills the full cell — its own green border replaces the delegate focus border
    editor->setGeometry(option.rect);
}

void CellDelegate::paint(QPainter* painter, const QStyleOptionViewItem& option,
                        const QModelIndex& index) const {
    painter->save();
    painter->setRenderHint(QPainter::Antialiasing, false);

    QRect rect = option.rect;
    bool isSelected = option.state & QStyle::State_Selected;
    bool hasFocus = option.state & QStyle::State_HasFocus;

    // --- Background ---
    QColor bgColor(Qt::white);
    QVariant bgData = index.data(Qt::BackgroundRole);
    if (bgData.isValid()) {
        QColor cellBg = bgData.value<QColor>();
        if (cellBg.isValid() && cellBg != QColor("#FFFFFF") && cellBg != QColor(Qt::white)) {
            bgColor = cellBg;
        }
    }

    if (isSelected && !hasFocus) {
        // Multi-select: light blue tint over cell background
        painter->fillRect(rect, bgColor);
        painter->fillRect(rect, QColor(198, 217, 240, 60));
    } else {
        painter->fillRect(rect, bgColor);
    }

    // --- Text ---
    QString text = index.data(Qt::DisplayRole).toString();
    if (!text.isEmpty()) {
        QFont font = option.font;
        QVariant fontData = index.data(Qt::FontRole);
        if (fontData.isValid()) {
            font = fontData.value<QFont>();
        }
        painter->setFont(font);

        QColor fgColor(Qt::black);
        QVariant fgData = index.data(Qt::ForegroundRole);
        if (fgData.isValid()) {
            QColor c = fgData.value<QColor>();
            if (c.isValid()) fgColor = c;
        }
        painter->setPen(fgColor);

        int alignment = Qt::AlignVCenter | Qt::AlignLeft;
        QVariant alignData = index.data(Qt::TextAlignmentRole);
        if (alignData.isValid()) {
            alignment = alignData.toInt();
        }

        // Indent support: add left padding per indent level
        int indentPx = 0;
        QVariant indentData = index.data(Qt::UserRole + 10);
        if (indentData.isValid()) {
            indentPx = indentData.toInt() * 12; // 12px per indent level
        }

        QRect textRect = rect.adjusted(4 + indentPx, 1, -4, -1);
        painter->drawText(textRect, alignment, text);
    }

    // --- Sparkline rendering ---
    QVariant sparkData = index.data(Qt::UserRole + 15); // SparklineRole
    if (sparkData.isValid() && sparkData.canConvert<SparklineRenderData>()) {
        auto rd = sparkData.value<SparklineRenderData>();
        if (!rd.values.isEmpty()) {
            QRect sparkRect = rect.adjusted(3, 3, -3, -3);
            drawSparkline(painter, sparkRect, rd);
        }
    }

    // --- Gridlines: single thin line on right and bottom edges ---
    if (m_showGridlines) {
        painter->setPen(QPen(QColor(218, 220, 224), 1, Qt::SolidLine));
        painter->drawLine(rect.right(), rect.top(), rect.right(), rect.bottom());
        painter->drawLine(rect.left(), rect.bottom(), rect.right(), rect.bottom());
    }

    // --- Cell borders (user-defined) ---
    auto drawBorder = [&](const QVariant& borderData, int x1, int y1, int x2, int y2) {
        if (!borderData.isValid()) return;
        QStringList parts = borderData.toString().split(',');
        if (parts.size() >= 2) {
            int w = parts[0].toInt();
            QColor c(parts[1]);
            if (c.isValid() && w > 0) {
                painter->setPen(QPen(c, w, Qt::SolidLine));
                painter->drawLine(x1, y1, x2, y2);
            }
        }
    };
    drawBorder(index.data(Qt::UserRole + 11), rect.left(), rect.top(), rect.right(), rect.top());      // top
    drawBorder(index.data(Qt::UserRole + 12), rect.left(), rect.bottom(), rect.right(), rect.bottom()); // bottom
    drawBorder(index.data(Qt::UserRole + 13), rect.left(), rect.top(), rect.left(), rect.bottom());     // left
    drawBorder(index.data(Qt::UserRole + 14), rect.right(), rect.top(), rect.right(), rect.bottom());   // right

    // --- Focus border: green rectangle for the current cell ---
    if (hasFocus) {
        painter->setClipRect(rect.adjusted(-1, -1, 1, 1)); // expand clip to avoid top/left clipping
        QPen focusPen(QColor(16, 124, 16), 2, Qt::SolidLine);
        painter->setPen(focusPen);
        painter->drawRect(rect.adjusted(1, 1, -1, -1));
    }

    painter->restore();
}

QSize CellDelegate::sizeHint(const QStyleOptionViewItem& option,
                             const QModelIndex& index) const {
    return QStyledItemDelegate::sizeHint(option, index);
}

void CellDelegate::drawSparkline(QPainter* painter, const QRect& rect,
                                  const SparklineRenderData& data) const {
    if (data.values.isEmpty() || rect.width() < 4 || rect.height() < 4) return;

    double range = data.maxVal - data.minVal;
    if (range == 0) range = 1.0;
    int n = data.values.size();

    switch (data.type) {
        case SparklineType::Line: {
            painter->setRenderHint(QPainter::Antialiasing, true);
            double stepX = static_cast<double>(rect.width()) / qMax(1, n - 1);

            QPainterPath path;
            for (int i = 0; i < n; ++i) {
                double x = rect.left() + i * stepX;
                double y = rect.bottom() - ((data.values[i] - data.minVal) / range) * rect.height();
                if (i == 0) path.moveTo(x, y);
                else path.lineTo(x, y);
            }
            painter->setPen(QPen(data.lineColor, data.lineWidth));
            painter->setBrush(Qt::NoBrush);
            painter->drawPath(path);

            if (data.showHighPoint && data.highIndex >= 0) {
                double x = rect.left() + data.highIndex * stepX;
                double y = rect.bottom() - ((data.values[data.highIndex] - data.minVal) / range) * rect.height();
                painter->setPen(Qt::NoPen);
                painter->setBrush(data.highPointColor);
                painter->drawEllipse(QPointF(x, y), 3, 3);
            }
            if (data.showLowPoint && data.lowIndex >= 0) {
                double x = rect.left() + data.lowIndex * stepX;
                double y = rect.bottom() - ((data.values[data.lowIndex] - data.minVal) / range) * rect.height();
                painter->setPen(Qt::NoPen);
                painter->setBrush(data.lowPointColor);
                painter->drawEllipse(QPointF(x, y), 3, 3);
            }
            painter->setRenderHint(QPainter::Antialiasing, false);
            break;
        }
        case SparklineType::Column: {
            double barW = static_cast<double>(rect.width()) / (n * 1.4);
            double zeroY = rect.bottom();
            if (data.minVal < 0) {
                zeroY = rect.bottom() - ((-data.minVal) / range) * rect.height();
            }
            for (int i = 0; i < n; ++i) {
                double x = rect.left() + i * (static_cast<double>(rect.width()) / n) +
                           (static_cast<double>(rect.width()) / n - barW) / 2;
                double barH = (std::abs(data.values[i]) / range) * rect.height();
                if (data.values[i] >= 0) {
                    QRectF bar(x, zeroY - barH, barW, barH);
                    painter->fillRect(bar, data.lineColor);
                } else {
                    QRectF bar(x, zeroY, barW, barH);
                    painter->fillRect(bar, data.negativeColor);
                }
            }
            break;
        }
        case SparklineType::WinLoss: {
            double barW = static_cast<double>(rect.width()) / (n * 1.4);
            double midY = rect.top() + rect.height() / 2.0;
            double halfH = rect.height() / 2.0 - 2;
            for (int i = 0; i < n; ++i) {
                double x = rect.left() + i * (static_cast<double>(rect.width()) / n) +
                           (static_cast<double>(rect.width()) / n - barW) / 2;
                if (data.values[i] > 0) {
                    QRectF bar(x, midY - halfH, barW, halfH);
                    painter->fillRect(bar, data.lineColor);
                } else if (data.values[i] < 0) {
                    QRectF bar(x, midY, barW, halfH);
                    painter->fillRect(bar, data.negativeColor);
                }
            }
            break;
        }
    }
}
