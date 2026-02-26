#include "CellDelegate.h"
#include "SpreadsheetView.h"
#include "FormulaPopupDelegate.h"
#include "Theme.h"
#include "../core/FormulaMetadata.h"
#include <QLineEdit>
#include <QPainter>
#include <QPainterPath>
#include <QStyleOptionViewItem>
#include <QApplication>
#include <QAbstractItemView>
#include <QListWidget>
#include <QKeyEvent>
#include <QTableView>
#include <QTimer>
#include <QLabel>
#include <QVBoxLayout>
#include <QFrame>
#include <QScrollBar>
#include <QScreen>
#include "../core/Spreadsheet.h"

// ===== Picklist tag color palettes (12 pastel bg + dark text pairs) =====
static const QColor s_tagBgColors[] = {
    QColor("#DBEAFE"), QColor("#FCE7F3"), QColor("#EDE9FE"), QColor("#D1FAE5"),
    QColor("#FEF3C7"), QColor("#FFE4E6"), QColor("#CFFAFE"), QColor("#FEE2E2"),
    QColor("#F3F4F6"), QColor("#ECFCCB"), QColor("#E0E7FF"), QColor("#FDF2F8")
};
static const QColor s_tagTextColors[] = {
    QColor("#1E40AF"), QColor("#9D174D"), QColor("#5B21B6"), QColor("#065F46"),
    QColor("#92400E"), QColor("#9F1239"), QColor("#155E75"), QColor("#991B1B"),
    QColor("#374151"), QColor("#3F6212"), QColor("#3730A3"), QColor("#831843")
};
static constexpr int TAG_COLOR_COUNT = 12;

QColor CellDelegate::tagBgColor(int index) { return s_tagBgColors[index % TAG_COLOR_COUNT]; }
QColor CellDelegate::tagTextColor(int index) { return s_tagTextColors[index % TAG_COLOR_COUNT]; }

// Extract the token currently being typed (after last delimiter)
static QString extractCurrentToken(const QString& text) {
    if (!text.startsWith("=") || text.length() <= 1) return {};
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
    return afterEq.mid(lastDelim + 1);
}

CellDelegate::CellDelegate(QObject* parent)
    : QStyledItemDelegate(parent) {
}

static void hideAllPopups(QObject* editorOrPopup) {
    // Retrieve editor from popup, or use directly
    QLineEdit* ed = qobject_cast<QLineEdit*>(editorOrPopup);
    if (!ed) ed = qobject_cast<QLineEdit*>(editorOrPopup->property("_editor").value<QObject*>());
    if (!ed) return;
    auto* p = qobject_cast<QListWidget*>(ed->property("_formulaPopup").value<QObject*>());
    auto* h = qobject_cast<QLabel*>(ed->property("_paramHint").value<QObject*>());
    auto* d = qobject_cast<QLabel*>(ed->property("_detailPanel").value<QObject*>());
    if (p) p->hide();
    if (h) h->hide();
    if (d) d->hide();
}

static void insertFunctionFromPopup(QListWidget* popup, QLineEdit* editor) {
    auto* item = popup->currentItem();
    if (!item) return;
    QString funcName = item->data(FuncNameRole).toString();
    hideAllPopups(editor);
    QString text = editor->text();
    QString token = extractCurrentToken(text);
    if (!token.isEmpty()) {
        int tokenStart = text.length() - token.length();
        editor->setText(text.left(tokenStart) + funcName + "(");
        editor->setCursorPosition(editor->text().length());
    }
    editor->setFocus();
}

bool CellDelegate::eventFilter(QObject* object, QEvent* event) {
    // Handle key events from the popup (Qt::Popup grabs keyboard)
    if (auto* popupWidget = qobject_cast<QListWidget*>(object)) {
        if (event->type() == QEvent::KeyPress) {
            auto* ke = static_cast<QKeyEvent*>(event);
            auto* ed = qobject_cast<QLineEdit*>(
                popupWidget->property("_editor").value<QObject*>());
            if (!ed) return false;

            if (ke->key() == Qt::Key_Down) {
                int next = popupWidget->currentRow() + 1;
                if (next < popupWidget->count()) popupWidget->setCurrentRow(next);
                return true;
            }
            if (ke->key() == Qt::Key_Up) {
                int prev = popupWidget->currentRow() - 1;
                if (prev >= 0) popupWidget->setCurrentRow(prev);
                return true;
            }
            if (ke->key() == Qt::Key_Return || ke->key() == Qt::Key_Enter ||
                ke->key() == Qt::Key_Tab) {
                insertFunctionFromPopup(popupWidget, ed);
                return true;
            }
            if (ke->key() == Qt::Key_Escape) {
                hideAllPopups(popupWidget);
                m_formulaEditMode = false;
                emit formulaEditModeChanged(false);
                emit closeEditor(ed, QAbstractItemDelegate::RevertModelCache);
                return true;
            }
            // Forward all other keys (typing) to the editor
            QApplication::sendEvent(ed, event);
            return true;
        }
        return false;
    }

    // During formula edit mode, block FocusOut from closing the editor
    if (m_formulaEditMode && event->type() == QEvent::FocusOut) {
        return true;
    }

    if (event->type() == QEvent::KeyPress) {
        QKeyEvent* keyEvent = static_cast<QKeyEvent*>(event);
        QLineEdit* editor = qobject_cast<QLineEdit*>(object);

        // Esc: cancel edit and revert value
        if (keyEvent->key() == Qt::Key_Escape && editor) {
            hideAllPopups(editor);
            m_formulaEditMode = false;
            emit formulaEditModeChanged(false);
            emit closeEditor(editor, QAbstractItemDelegate::RevertModelCache);
            return true;
        }

        // Arrow keys during editing: commit and move (like Excel)
        // But NOT during formula edit mode — arrows navigate in the formula text
        if (!m_formulaEditMode &&
            (keyEvent->key() == Qt::Key_Up || keyEvent->key() == Qt::Key_Down ||
             keyEvent->key() == Qt::Key_Left || keyEvent->key() == Qt::Key_Right)) {
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
    // Checkbox cells: no editor — toggle is handled in SpreadsheetView::mousePressEvent
    if (m_spreadsheetView) {
        auto spreadsheet = m_spreadsheetView->getSpreadsheet();
        if (spreadsheet) {
            auto cell = spreadsheet->getCellIfExists(index.row(), index.column());
            if (cell) {
                const auto& style = cell->getStyle();
                if (style.numberFormat == "Checkbox") return nullptr;
                // Picklist cells: handled by SpreadsheetView popup
                if (style.numberFormat == "Picklist") return nullptr;
            }
        }
    }

    QLineEdit* editor = new QLineEdit(parent);
    editor->setFrame(false);
    editor->setStyleSheet(QString(
        "QLineEdit { background: white; padding: 1px 2px; "
        "border: 2px solid %1; selection-background-color: #0078D4; }")
        .arg(ThemeManager::instance().currentTheme().editorBorderColor.name()));

    // --- Custom formula autocomplete popup ---
    auto* popup = new QListWidget(parent->window());
    popup->setWindowFlags(Qt::Popup | Qt::FramelessWindowHint);
    popup->setAttribute(Qt::WA_ShowWithoutActivating);
    popup->setFocusPolicy(Qt::NoFocus);
    popup->setHorizontalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    popup->setVerticalScrollBarPolicy(Qt::ScrollBarAsNeeded);
    popup->setStyleSheet(
        "QListWidget { background: white; border: 1px solid #C0C0C0; outline: none; "
        "border-radius: 6px; }"
        "QListWidget::item { padding: 0px; border: none; }"
        "QListWidget::item:selected { background: transparent; }"
        "QListWidget::item:hover { background: transparent; }"
    );
    popup->setItemDelegate(new FormulaPopupDelegate(popup));
    popup->hide();

    // --- Parameter hint tooltip ---
    auto* paramHint = new QLabel(parent->window());
    paramHint->setWindowFlags(Qt::ToolTip | Qt::FramelessWindowHint);
    paramHint->setAttribute(Qt::WA_ShowWithoutActivating);
    paramHint->setStyleSheet(
        "QLabel { background: #FFF8DC; border: 1px solid #E0D8B0; padding: 4px 8px; "
        "font-size: 12px; color: #333; border-radius: 3px; }");
    paramHint->setTextFormat(Qt::RichText);
    paramHint->hide();

    // --- Detail panel (shown on click) ---
    auto* detailPanel = new QLabel(parent->window());
    detailPanel->setWindowFlags(Qt::ToolTip | Qt::FramelessWindowHint);
    detailPanel->setAttribute(Qt::WA_ShowWithoutActivating);
    detailPanel->setStyleSheet(
        "QLabel { background: white; border: 1px solid #D0D0D0; padding: 12px; "
        "border-radius: 6px; }");
    detailPanel->setTextFormat(Qt::RichText);
    detailPanel->setWordWrap(true);
    detailPanel->setFixedWidth(340);
    detailPanel->hide();

    // Populate popup items from registry (startsWith filter)
    auto populatePopup = [popup](const QString& prefix) {
        popup->clear();
        const auto& reg = formulaRegistry();
        for (auto it = reg.begin(); it != reg.end(); ++it) {
            if (it->name.startsWith(prefix, Qt::CaseInsensitive)) {
                auto* item = new QListWidgetItem(popup);
                item->setData(FuncNameRole, it->name);
                item->setData(FuncDescRole, it->description);
            }
        }
        return popup->count() > 0;
    };

    // When item clicked in popup → show detail panel
    QObject::connect(popup, &QListWidget::itemClicked, editor,
        [popup, detailPanel](QListWidgetItem* item) {
            QString funcName = item->data(FuncNameRole).toString();
            const auto& reg = formulaRegistry();
            if (reg.contains(funcName)) {
                detailPanel->setText(buildDetailHtml(reg[funcName]));
                detailPanel->adjustSize();
                // Position to the right of popup
                QPoint pos = popup->mapToGlobal(QPoint(popup->width() + 4, 0));
                detailPanel->move(pos);
                detailPanel->show();
            }
        });

    // Install event filter on popup (to forward keys to editor) and on editor
    popup->installEventFilter(const_cast<CellDelegate*>(this));
    popup->setProperty("_editor", QVariant::fromValue(static_cast<QObject*>(editor)));
    editor->installEventFilter(const_cast<CellDelegate*>(this));

    // Store widgets on editor for access in eventFilter
    editor->setProperty("_formulaPopup", QVariant::fromValue(static_cast<QObject*>(popup)));
    editor->setProperty("_paramHint", QVariant::fromValue(static_cast<QObject*>(paramHint)));
    editor->setProperty("_detailPanel", QVariant::fromValue(static_cast<QObject*>(detailPanel)));

    // Show popup / param hint on text change
    QObject::connect(editor, &QLineEdit::textChanged, editor,
        [this, editor, popup, paramHint, detailPanel, populatePopup](const QString& text) {
        emit formulaEditModeChanged(text.startsWith("="));
        detailPanel->hide(); // hide detail panel on any text change

        if (text.startsWith("=") && text.length() > 1) {
            QString token = extractCurrentToken(text);

            // Show autocomplete if typing a function name token
            if (!token.isEmpty() && token[0].isLetter()) {
                if (populatePopup(token)) {
                    popup->setCurrentRow(0);
                    // Defer positioning to next event loop — editor geometry
                    // may not be finalized on the very first keystroke
                    QTimer::singleShot(0, editor, [editor, popup]() {
                        if (!editor || !popup) return;
                        int popupW = qMax(460, editor->width());
                        int visibleItems = qMin(popup->count(), 8);
                        int popupH = visibleItems * 30 + 6;
                        popup->setFixedWidth(popupW);
                        popup->setFixedHeight(popupH);
                        int edH = qMax(editor->height(), 25);
                        QPoint below = editor->mapToGlobal(QPoint(0, edH + 2));
                        QPoint above = editor->mapToGlobal(QPoint(0, -popupH - 2));
                        QScreen* screen = editor->screen();
                        if (screen && below.y() + popupH > screen->availableGeometry().bottom()) {
                            popup->move(above);
                        } else {
                            popup->move(below);
                        }
                        popup->show();
                    });
                } else {
                    popup->hide();
                }
            } else {
                popup->hide();
            }

            // Show param hint if inside a function call
            int cursorPos = editor->cursorPosition();
            FormulaContext ctx = findFormulaContext(text, cursorPos);
            const auto& reg = formulaRegistry();
            if (ctx.paramIndex >= 0 && reg.contains(ctx.funcName)) {
                const auto& info = reg[ctx.funcName];
                paramHint->setText(buildParamHintHtml(info, ctx.paramIndex));
                paramHint->adjustSize();
                // Defer positioning to next event loop
                QTimer::singleShot(0, editor, [editor, popup, paramHint]() {
                    if (!editor || !paramHint) return;
                    QPoint hintPos;
                    if (popup && popup->isVisible()) {
                        hintPos = popup->mapToGlobal(QPoint(0, popup->height() + 2));
                    } else {
                        int edH = qMax(editor->height(), 25);
                        hintPos = editor->mapToGlobal(QPoint(0, edH + 2));
                    }
                    paramHint->move(hintPos);
                    paramHint->show();
                });
            } else {
                paramHint->hide();
            }
        } else {
            popup->hide();
            paramHint->hide();
        }
    });

    // Also update param hint on cursor position change
    QObject::connect(editor, &QLineEdit::cursorPositionChanged, editor,
        [editor, paramHint, popup](int, int newPos) {
        QString text = editor->text();
        if (!text.startsWith("=")) { paramHint->hide(); return; }
        FormulaContext ctx = findFormulaContext(text, newPos);
        const auto& reg = formulaRegistry();
        if (ctx.paramIndex >= 0 && reg.contains(ctx.funcName)) {
            const auto& info = reg[ctx.funcName];
            paramHint->setText(buildParamHintHtml(info, ctx.paramIndex));
            paramHint->adjustSize();
            QPoint hintPos;
            if (popup && popup->isVisible()) {
                hintPos = popup->mapToGlobal(QPoint(0, popup->height() + 2));
            } else {
                int edH = qMax(editor->height(), 25);
                hintPos = editor->mapToGlobal(QPoint(0, edH + 2));
            }
            paramHint->move(hintPos);
            paramHint->show();
        } else {
            paramHint->hide();
        }
    });

    // Clean up on editor destruction — hide immediately, then delete
    QObject::connect(editor, &QObject::destroyed, popup, [popup]() {
        popup->hide();
        popup->deleteLater();
    });
    QObject::connect(editor, &QObject::destroyed, paramHint, [paramHint]() {
        paramHint->hide();
        paramHint->deleteLater();
    });
    QObject::connect(editor, &QObject::destroyed, detailPanel, [detailPanel]() {
        detailPanel->hide();
        detailPanel->deleteLater();
    });

    return editor;
}

void CellDelegate::setEditorData(QWidget* editor, const QModelIndex& index) const {
    QLineEdit* lineEdit = qobject_cast<QLineEdit*>(editor);
    if (lineEdit) {
        lineEdit->setText(index.data(Qt::EditRole).toString());
        // Place cursor at end (not selecting all text) — like Excel
        // Deferred to after Qt finishes initializing the editor, which otherwise re-selects all text
        QTimer::singleShot(0, lineEdit, [lineEdit]() {
            lineEdit->deselect();
            lineEdit->setCursorPosition(lineEdit->text().length());
        });
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
        if (cellBg.isValid() && cellBg.rgb() != QColor(Qt::white).rgb()) {
            bgColor = cellBg;
        }
    }

    if (isSelected && !hasFocus) {
        // Multi-select: light blue tint over cell background
        painter->fillRect(rect, bgColor);
        painter->fillRect(rect, ThemeManager::instance().currentTheme().selectionTint);
    } else {
        painter->fillRect(rect, bgColor);
    }

    // --- Formula recalc flash overlay (bright yellow, fade-in then hold then fade-out) ---
    if (m_spreadsheetView) {
        double flashProgress = m_spreadsheetView->cellAnimationProgress(index.row(), index.column());
        if (flashProgress > 0.01) {
            int alpha = static_cast<int>(130 * flashProgress);
            painter->fillRect(rect, QColor(255, 243, 100, alpha));
        }
    }

    // --- Picklist / Checkbox rendering ---
    bool skipText = false;
    if (m_spreadsheetView) {
        auto spreadsheet = m_spreadsheetView->getSpreadsheet();
        if (spreadsheet) {
            auto cell = spreadsheet->getCellIfExists(index.row(), index.column());
            if (cell) {
                const auto& style = cell->getStyle();
                if (style.numberFormat == "Checkbox") {
                    bool checked = false;
                    auto val = cell->getValue();
                    if (val.typeId() == QMetaType::Bool) checked = val.toBool();
                    else {
                        QString s = val.toString().toLower();
                        checked = (s == "true" || s == "1");
                    }
                    drawCheckbox(painter, rect, checked);
                    skipText = true;
                } else if (style.numberFormat == "Picklist") {
                    QString val = cell->getValue().toString();
                    // Always fill white bg to override focus/selection highlight
                    painter->fillRect(rect, Qt::white);
                    const auto* rule = spreadsheet->getValidationAt(index.row(), index.column());
                    QStringList options = rule ? rule->listItems : QStringList();
                    QStringList colors = rule ? rule->listItemColors : QStringList();
                    Qt::Alignment align = Qt::AlignLeft | Qt::AlignVCenter;
                    QVariant alignVar = index.data(Qt::TextAlignmentRole);
                    if (alignVar.isValid())
                        align = Qt::Alignment(alignVar.toInt());
                    drawPicklistTags(painter, rect, val, options, colors, align);
                    skipText = true;
                }
            }
        }
    }

    // --- Text ---
    QString text = index.data(Qt::DisplayRole).toString();
    if (!skipText && !text.isEmpty()) {
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

        // Text rotation support
        int rotation = 0;
        QVariant rotData = index.data(Qt::UserRole + 16);
        if (rotData.isValid()) {
            rotation = rotData.toInt();
        }

        QRect textRect = rect.adjusted(4 + indentPx, 1, -4, -1);

        if (rotation == 0) {
            painter->drawText(textRect, alignment, text);
        } else if (rotation == 270) {
            // Vertical stacked text: draw each character on its own line
            QFontMetrics fm(font);
            int charH = fm.height();
            int maxCharW = 0;
            for (const QChar& ch : text) {
                maxCharW = qMax(maxCharW, fm.horizontalAdvance(ch));
            }
            int totalH = charH * text.length();
            int startY = textRect.top() + qMax(0, (textRect.height() - totalH) / 2);
            int centerX = textRect.left() + (textRect.width() - maxCharW) / 2;
            for (int i = 0; i < text.length(); ++i) {
                QRect charRect(centerX, startY + i * charH, maxCharW, charH);
                painter->drawText(charRect, Qt::AlignCenter, QString(text[i]));
            }
        } else {
            // Angled text: rotate painter
            painter->save();
            QFontMetrics fm(font);
            int textW = fm.horizontalAdvance(text);
            int textH = fm.height();
            QPointF center = textRect.center();
            painter->translate(center);
            painter->rotate(-rotation); // Qt rotates clockwise, we want CCW for positive angles
            QRectF rotRect(-textW / 2.0, -textH / 2.0, textW, textH);
            painter->drawText(rotRect, Qt::AlignCenter, text);
            painter->restore();
        }
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
        painter->setPen(QPen(ThemeManager::instance().currentTheme().gridLineColor, 1, Qt::SolidLine));
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
            int ps = (parts.size() >= 3) ? parts[2].toInt() : 0;
            if (c.isValid() && w > 0) {
                Qt::PenStyle penStyle = Qt::SolidLine;
                if (ps == 1) penStyle = Qt::DashLine;
                else if (ps == 2) penStyle = Qt::DotLine;
                painter->setPen(QPen(c, w, penStyle));
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
        QPen focusPen(ThemeManager::instance().currentTheme().focusBorderColor, 2, Qt::SolidLine);
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

void CellDelegate::drawCheckbox(QPainter* painter, const QRect& rect, bool checked) const {
    int boxSize = 14;
    int x = rect.left() + (rect.width() - boxSize) / 2;
    int y = rect.top() + (rect.height() - boxSize) / 2;
    QRectF boxRect(x + 0.5, y + 0.5, boxSize - 1, boxSize - 1);

    painter->setRenderHint(QPainter::Antialiasing, true);
    if (checked) {
        // Soft green filled rounded rect
        painter->setPen(Qt::NoPen);
        painter->setBrush(ThemeManager::instance().currentTheme().checkboxChecked);
        painter->drawRoundedRect(boxRect, 4, 4);
        // Smooth checkmark
        QPainterPath check;
        check.moveTo(x + 3.5, y + 7);
        check.lineTo(x + 6, y + 10);
        check.lineTo(x + 10.5, y + 4.5);
        painter->setPen(QPen(Qt::white, 1.6, Qt::SolidLine, Qt::RoundCap, Qt::RoundJoin));
        painter->setBrush(Qt::NoBrush);
        painter->drawPath(check);
    } else {
        // Soft outlined rounded rect
        painter->setPen(QPen(ThemeManager::instance().currentTheme().checkboxUncheckedBorder, 1.2));
        painter->setBrush(QColor("#FAFAFA"));
        painter->drawRoundedRect(boxRect, 4, 4);
    }
    painter->setRenderHint(QPainter::Antialiasing, false);
}

void CellDelegate::drawPicklistTags(QPainter* painter, const QRect& rect,
                                     const QString& value, const QStringList& allOptions,
                                     const QStringList& optionColors,
                                     Qt::Alignment alignment) const {
    painter->setRenderHint(QPainter::Antialiasing, true);

    // Draw a subtle dropdown arrow on the right
    {
        int arrowSize = 6;
        int arrowX = rect.right() - arrowSize - 6;
        int arrowY = rect.top() + (rect.height() - arrowSize / 2) / 2;
        QPainterPath arrow;
        arrow.moveTo(arrowX, arrowY);
        arrow.lineTo(arrowX + arrowSize, arrowY);
        arrow.lineTo(arrowX + arrowSize / 2.0, arrowY + arrowSize * 0.55);
        arrow.closeSubpath();
        painter->setPen(Qt::NoPen);
        painter->setBrush(QColor("#B0B4BA"));
        painter->drawPath(arrow);
    }

    if (value.isEmpty()) {
        painter->setRenderHint(QPainter::Antialiasing, false);
        return;
    }

    QStringList selected = value.split('|', Qt::SkipEmptyParts);

    QFont tagFont(painter->font().family(), 10);
    tagFont.setWeight(QFont::Medium);
    tagFont.setLetterSpacing(QFont::AbsoluteSpacing, 0.2);
    QFontMetrics fm(tagFont);
    painter->setFont(tagFont);

    int tagH = 18;
    int gap = 3;
    int maxX = rect.right() - 18; // leave room for dropdown arrow

    // Compute total width for horizontal alignment
    int totalTagW = 0;
    for (const QString& item : selected) {
        QString trimmed = item.trimmed();
        if (!trimmed.isEmpty())
            totalTagW += fm.horizontalAdvance(trimmed) + 14 + gap;
    }
    if (totalTagW > 0) totalTagW -= gap;

    int availW = maxX - rect.left() - 5;
    int x = rect.left() + 5;
    if (alignment & Qt::AlignHCenter)
        x = rect.left() + 5 + qMax(0, (availW - totalTagW) / 2);
    else if (alignment & Qt::AlignRight)
        x = rect.left() + 5 + qMax(0, availW - totalTagW);

    int y;
    if (alignment & Qt::AlignTop)
        y = rect.top() + 2;
    else if (alignment & Qt::AlignBottom)
        y = rect.bottom() - tagH - 2;
    else
        y = rect.top() + (rect.height() - tagH) / 2;

    int tagIndex = 0;
    for (const QString& item : selected) {
        QString trimmed = item.trimmed();
        if (trimmed.isEmpty()) continue;

        int colorIdx = allOptions.indexOf(trimmed);
        if (colorIdx < 0) colorIdx = static_cast<int>(qHash(trimmed) % TAG_COLOR_COUNT);

        QColor bg, fg;
        // Use custom color if provided for this option
        if (colorIdx >= 0 && colorIdx < optionColors.size() && !optionColors[colorIdx].isEmpty()) {
            bg = QColor(optionColors[colorIdx]);
            // Compute readable text color: dark text on light bg, white on dark bg
            fg = (bg.lightness() > 140) ? bg.darker(300) : QColor(Qt::white);
        } else {
            bg = tagBgColor(colorIdx);
            fg = tagTextColor(colorIdx);
        }

        int textW = fm.horizontalAdvance(trimmed);
        int tagW = textW + 14;

        // If not the first tag and doesn't fit, show "..." and stop
        if (tagIndex > 0 && x + tagW > maxX) {
            painter->setPen(QColor("#9CA3AF"));
            QFont smallFont(painter->font());
            smallFont.setPointSize(8);
            painter->setFont(smallFont);
            painter->drawText(QRectF(x, y, 20, tagH), Qt::AlignVCenter | Qt::AlignLeft, "...");
            break;
        }

        // For the first tag, clip to available width if needed
        int drawW = qMin(tagW, maxX - x);
        QRectF tagRect(x, y, drawW, tagH);
        painter->setPen(Qt::NoPen);
        painter->setBrush(bg);
        painter->drawRoundedRect(tagRect, tagH / 2.0, tagH / 2.0);

        painter->setPen(fg);
        painter->save();
        painter->setClipRect(tagRect);
        painter->drawText(QRectF(x, y, tagW, tagH), Qt::AlignCenter, trimmed);
        painter->restore();

        x += drawW + gap;
        tagIndex++;
    }

    painter->setRenderHint(QPainter::Antialiasing, false);
}
