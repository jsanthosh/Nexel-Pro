#include "ShapeWidget.h"
#include "Theme.h"
#include "MainWindow.h"
#include <QPainter>
#include <QPainterPath>
#include <QMouseEvent>
#include <QKeyEvent>
#include <QMenu>
#include <QInputDialog>
#include <QApplication>
#include <QScrollBar>
#include <QAbstractScrollArea>
#include <QtMath>
#include <cmath>

ShapeWidget::ShapeWidget(QWidget* parent)
    : QWidget(parent) {
    setMinimumSize(30, 30);
    resize(150, 100);
    setAttribute(Qt::WA_DeleteOnClose, false);
    setMouseTracking(true);
    setFocusPolicy(Qt::ClickFocus);

    // Auto-scroll timer for dragging near viewport edges
    m_autoScrollTimer = new QTimer(this);
    m_autoScrollTimer->setInterval(16);  // ~60fps for fast, smooth scrolling
    connect(m_autoScrollTimer, &QTimer::timeout, this, &ShapeWidget::onAutoScroll);
}

void ShapeWidget::setConfig(const ShapeConfig& config) {
    m_config = config;
    update();
}

void ShapeWidget::setSelected(bool selected) {
    m_selected = selected;
    update();
}

// --- Paint ---

void ShapeWidget::paintEvent(QPaintEvent*) {
    QPainter p(this);
    p.setRenderHint(QPainter::Antialiasing, true);

    QRect area = rect().adjusted(HANDLE_SIZE / 2 + 1, HANDLE_SIZE / 2 + 1,
                                 -HANDLE_SIZE / 2 - 1, -HANDLE_SIZE / 2 - 1);
    drawShape(p, area);

    // Draw text if any
    if (!m_config.text.isEmpty()) {
        p.setPen(m_config.textColor);
        p.setFont(QFont("Arial", m_config.fontSize));
        p.drawText(area, Qt::AlignCenter | Qt::TextWordWrap, m_config.text);
    }

    if (m_selected) drawSelectionHandles(p);
}

void ShapeWidget::drawShape(QPainter& p, const QRect& area) {
    p.setOpacity(m_config.opacity);
    p.setPen(QPen(m_config.strokeColor, m_config.strokeWidth));
    p.setBrush(m_config.fillColor);

    switch (m_config.type) {
        case ShapeType::Rectangle:    drawRectangle(p, area); break;
        case ShapeType::RoundedRect:  drawRoundedRect(p, area); break;
        case ShapeType::Circle:       drawCircle(p, area); break;
        case ShapeType::Ellipse:      drawEllipse(p, area); break;
        case ShapeType::Triangle:     drawTriangle(p, area); break;
        case ShapeType::Star:         drawStar(p, area); break;
        case ShapeType::Arrow:        drawArrow(p, area); break;
        case ShapeType::Line:         drawShapeLine(p, area); break;
        case ShapeType::Diamond:      drawDiamond(p, area); break;
        case ShapeType::Pentagon:     drawPentagon(p, area); break;
        case ShapeType::Hexagon:      drawHexagon(p, area); break;
        case ShapeType::Callout:      drawCallout(p, area); break;
    }

    p.setOpacity(1.0);
}

void ShapeWidget::drawRectangle(QPainter& p, const QRect& area) {
    p.drawRect(area);
}

void ShapeWidget::drawRoundedRect(QPainter& p, const QRect& area) {
    float r = m_config.cornerRadius > 0 ? m_config.cornerRadius : 10;
    p.drawRoundedRect(area, r, r);
}

void ShapeWidget::drawCircle(QPainter& p, const QRect& area) {
    int size = qMin(area.width(), area.height());
    QRect sq(area.center().x() - size / 2, area.center().y() - size / 2, size, size);
    p.drawEllipse(sq);
}

void ShapeWidget::drawEllipse(QPainter& p, const QRect& area) {
    p.drawEllipse(area);
}

void ShapeWidget::drawTriangle(QPainter& p, const QRect& area) {
    QPolygon tri;
    tri << QPoint(area.center().x(), area.top())
        << QPoint(area.left(), area.bottom())
        << QPoint(area.right(), area.bottom());
    p.drawPolygon(tri);
}

void ShapeWidget::drawStar(QPainter& p, const QRect& area) {
    int cx = area.center().x(), cy = area.center().y();
    int outerR = qMin(area.width(), area.height()) / 2;
    int innerR = outerR * 40 / 100;

    QPolygonF star;
    for (int i = 0; i < 10; ++i) {
        double angle = M_PI / 2 + i * M_PI / 5;
        int r = (i % 2 == 0) ? outerR : innerR;
        star << QPointF(cx + r * std::cos(angle), cy - r * std::sin(angle));
    }
    p.drawPolygon(star);
}

void ShapeWidget::drawArrow(QPainter& p, const QRect& area) {
    int headW = area.width() / 3;
    int shaftH = area.height() / 3;

    QPolygon arrow;
    arrow << QPoint(area.right(), area.center().y())                   // tip
          << QPoint(area.right() - headW, area.top())                  // top of head
          << QPoint(area.right() - headW, area.center().y() - shaftH / 2)
          << QPoint(area.left(), area.center().y() - shaftH / 2)      // shaft top
          << QPoint(area.left(), area.center().y() + shaftH / 2)      // shaft bottom
          << QPoint(area.right() - headW, area.center().y() + shaftH / 2)
          << QPoint(area.right() - headW, area.bottom());             // bottom of head
    p.drawPolygon(arrow);
}

void ShapeWidget::drawDiamond(QPainter& p, const QRect& area) {
    QPolygon diamond;
    diamond << QPoint(area.center().x(), area.top())
            << QPoint(area.right(), area.center().y())
            << QPoint(area.center().x(), area.bottom())
            << QPoint(area.left(), area.center().y());
    p.drawPolygon(diamond);
}

void ShapeWidget::drawPentagon(QPainter& p, const QRect& area) {
    int cx = area.center().x(), cy = area.center().y();
    int r = qMin(area.width(), area.height()) / 2;
    QPolygonF poly;
    for (int i = 0; i < 5; ++i) {
        double angle = M_PI / 2 + i * 2 * M_PI / 5;
        poly << QPointF(cx + r * std::cos(angle), cy - r * std::sin(angle));
    }
    p.drawPolygon(poly);
}

void ShapeWidget::drawHexagon(QPainter& p, const QRect& area) {
    int cx = area.center().x(), cy = area.center().y();
    int r = qMin(area.width(), area.height()) / 2;
    QPolygonF poly;
    for (int i = 0; i < 6; ++i) {
        double angle = i * M_PI / 3;
        poly << QPointF(cx + r * std::cos(angle), cy - r * std::sin(angle));
    }
    p.drawPolygon(poly);
}

void ShapeWidget::drawCallout(QPainter& p, const QRect& area) {
    QPainterPath path;
    QRect box(area.left(), area.top(), area.width(), area.height() - 15);
    path.addRoundedRect(box, 8, 8);

    // Tail
    int tailX = box.left() + box.width() / 4;
    path.moveTo(tailX, box.bottom());
    path.lineTo(tailX - 5, area.bottom());
    path.lineTo(tailX + 15, box.bottom());

    p.drawPath(path);
}

void ShapeWidget::drawShapeLine(QPainter& p, const QRect& area) {
    p.setBrush(Qt::NoBrush);
    p.drawLine(area.topLeft(), area.bottomRight());
}

void ShapeWidget::drawSelectionHandles(QPainter& p) {
    QColor handleColor = ThemeManager::instance().currentTheme().selectionHandleColor;
    p.setPen(QPen(handleColor, 2));
    p.setBrush(Qt::NoBrush);
    p.drawRect(rect().adjusted(1, 1, -2, -2));

    p.setPen(QPen(handleColor, 1));
    p.setBrush(Qt::white);

    auto drawHandle = [&](int cx, int cy) {
        p.drawRect(cx - HANDLE_SIZE / 2, cy - HANDLE_SIZE / 2, HANDLE_SIZE, HANDLE_SIZE);
    };

    int w = width(), h = height();
    drawHandle(0, 0);
    drawHandle(w - 1, 0);
    drawHandle(0, h - 1);
    drawHandle(w - 1, h - 1);
    drawHandle(w / 2, 0);
    drawHandle(w / 2, h - 1);
    drawHandle(0, h / 2);
    drawHandle(w - 1, h / 2);
}

// --- Mouse interaction ---

ShapeWidget::ResizeHandle ShapeWidget::hitTestHandle(const QPoint& pos) const {
    if (!m_selected) return None;

    int w = width(), h = height();
    int hs = HANDLE_SIZE;

    if (QRect(-hs / 2, -hs / 2, hs, hs).contains(pos)) return TopLeft;
    if (QRect(w - hs / 2, -hs / 2, hs, hs).contains(pos)) return TopRight;
    if (QRect(-hs / 2, h - hs / 2, hs, hs).contains(pos)) return BottomLeft;
    if (QRect(w - hs / 2, h - hs / 2, hs, hs).contains(pos)) return BottomRight;
    if (QRect(w / 2 - hs / 2, -hs / 2, hs, hs).contains(pos)) return Top;
    if (QRect(w / 2 - hs / 2, h - hs / 2, hs, hs).contains(pos)) return Bottom;
    if (QRect(-hs / 2, h / 2 - hs / 2, hs, hs).contains(pos)) return Left;
    if (QRect(w - hs / 2, h / 2 - hs / 2, hs, hs).contains(pos)) return Right;

    return None;
}

void ShapeWidget::mousePressEvent(QMouseEvent* event) {
    if (event->button() == Qt::LeftButton) {
        setSelected(true);
        emit shapeSelected(this);

        ResizeHandle handle = hitTestHandle(event->pos());
        if (handle != None) {
            m_resizing = true;
            m_activeHandle = handle;
            m_dragStart = event->globalPosition().toPoint();
            m_resizeStartGeometry = geometry();
        } else {
            m_dragging = true;
            m_dragStart = event->globalPosition().toPoint();
            m_dragOffset = m_dragStart - pos();
        }
    }
}

void ShapeWidget::mouseMoveEvent(QMouseEvent* event) {
    if (m_resizing) {
        QPoint delta = event->globalPosition().toPoint() - m_dragStart;
        QRect geo = m_resizeStartGeometry;

        switch (m_activeHandle) {
            case TopLeft:     geo.setTopLeft(geo.topLeft() + delta); break;
            case TopRight:    geo.setTopRight(geo.topRight() + delta); break;
            case BottomLeft:  geo.setBottomLeft(geo.bottomLeft() + delta); break;
            case BottomRight: geo.setBottomRight(geo.bottomRight() + delta); break;
            case Top:         geo.setTop(geo.top() + delta.y()); break;
            case Bottom:      geo.setBottom(geo.bottom() + delta.y()); break;
            case Left:        geo.setLeft(geo.left() + delta.x()); break;
            case Right:       geo.setRight(geo.right() + delta.x()); break;
            default: break;
        }

        if (geo.width() >= minimumWidth() && geo.height() >= minimumHeight()) {
            setGeometry(geo);
        }
    } else if (m_dragging) {
        QPoint globalPos = event->globalPosition().toPoint();
        m_lastGlobalPos = globalPos;
        QPoint rawPos = globalPos - m_dragOffset;
        QPoint newPos = rawPos;
        // Constrain to parent viewport
        bool clampedAtEdge = false;
        if (parentWidget()) {
            int maxX = parentWidget()->width() - width();
            int maxY = parentWidget()->height() - height();
            newPos.setX(qBound(0, rawPos.x(), maxX));
            newPos.setY(qBound(0, rawPos.y(), maxY));
            // Auto-scroll triggers when shape hits any viewport edge
            clampedAtEdge = (rawPos.x() < 0 || rawPos.x() > maxX ||
                             rawPos.y() < 0 || rawPos.y() > maxY);
        }
        QPoint delta = newPos - pos();
        move(newPos);
        // Group-aware drag: move siblings by same delta
        int gid = property("overlayGroupId").toInt();
        if (gid > 0 && parentWidget()) {
            for (QWidget* sibling : parentWidget()->findChildren<QWidget*>()) {
                if (sibling != this && sibling->property("overlayGroupId").toInt() == gid)
                    sibling->move(sibling->pos() + delta);
            }
        }
        emit shapeMoved(this);

        // Start auto-scroll as soon as shape hits viewport edge.
        // Don't stop here on !clampedAtEdge — let onAutoScroll() handle stopping
        // so a tiny mouse wiggle doesn't kill the scroll.
        if (clampedAtEdge && !m_autoScrollTimer->isActive())
            m_autoScrollTimer->start();
    } else if (m_selected) {
        ResizeHandle h = hitTestHandle(event->pos());
        switch (h) {
            case TopLeft: case BottomRight: setCursor(Qt::SizeFDiagCursor); break;
            case TopRight: case BottomLeft: setCursor(Qt::SizeBDiagCursor); break;
            case Top: case Bottom: setCursor(Qt::SizeVerCursor); break;
            case Left: case Right: setCursor(Qt::SizeHorCursor); break;
            default: setCursor(Qt::SizeAllCursor); break;
        }
    }
}

void ShapeWidget::mouseReleaseEvent(QMouseEvent*) {
    m_dragging = false;
    m_resizing = false;
    m_activeHandle = None;
    m_autoScrollTimer->stop();
}

void ShapeWidget::onAutoScroll() {
    if (!m_dragging || !parentWidget()) return;

    // parentWidget() is the viewport; its parent is the QAbstractScrollArea (QTableView)
    auto* scrollArea = qobject_cast<QAbstractScrollArea*>(parentWidget()->parentWidget());
    if (!scrollArea) return;

    // Determine scroll direction from where the unclamped position would be
    QPoint rawPos = m_lastGlobalPos - m_dragOffset;
    int maxX = parentWidget()->width() - width();
    int maxY = parentWidget()->height() - height();

    int dx = 0, dy = 0;
    if (rawPos.y() > maxY)
        dy = 5;
    else if (rawPos.y() < 0)
        dy = -5;
    if (rawPos.x() > maxX)
        dx = 2;
    else if (rawPos.x() < 0)
        dx = -2;

    if (dx == 0 && dy == 0) { m_autoScrollTimer->stop(); return; }

    QScrollBar* vBar = scrollArea->verticalScrollBar();
    QScrollBar* hBar = scrollArea->horizontalScrollBar();

    int oldV = vBar->value(), oldH = hBar->value();
    if (dy) vBar->setValue(vBar->value() + dy);
    if (dx) hBar->setValue(hBar->value() + dx);

    // Stop if scroll didn't actually happen (hit min/max)
    if (vBar->value() == oldV && hBar->value() == oldH) {
        m_autoScrollTimer->stop();
        return;
    }

    emit shapeMoved(this);
}

void ShapeWidget::mouseDoubleClickEvent(QMouseEvent*) {
    // Allow editing text in shape
    bool ok;
    QString text = QInputDialog::getText(this, "Shape Text",
        "Enter text for shape:", QLineEdit::Normal, m_config.text, &ok);
    if (ok) {
        m_config.text = text;
        update();
    }
}

void ShapeWidget::contextMenuEvent(QContextMenuEvent* event) {
    QMenu menu(this);
    {
        const auto& t = ThemeManager::instance().currentTheme();
        menu.setStyleSheet(QString(
            "QMenu { background: %1; border: 1px solid %2; border-radius: 6px; padding: 4px; }"
            "QMenu::item { padding: 6px 20px; border-radius: 4px; }"
            "QMenu::item:selected { background-color: %3; }"
            "QMenu::separator { height: 1px; background: %2; margin: 3px 8px; }")
            .arg(t.popupBackground.name(), t.popupBorder.name(), t.popupItemSelected.name()));
    }

    menu.addAction("Edit Text...", this, [this]() {
        bool ok;
        QString text = QInputDialog::getText(this, "Shape Text",
            "Enter text:", QLineEdit::Normal, m_config.text, &ok);
        if (ok) { m_config.text = text; update(); }
    });
    menu.addAction("Edit Shape...", this, [this]() { emit editRequested(this); });
    menu.addSeparator();

    // Order submenu
    QMenu* orderMenu = menu.addMenu("Order");
    orderMenu->setStyleSheet(menu.styleSheet());
    auto* mw = qobject_cast<MainWindow*>(window());
    if (mw) {
        orderMenu->addAction("Bring to Front", this, [mw, this]() { mw->bringToFront(this); });
        orderMenu->addAction("Send to Back", this, [mw, this]() { mw->sendToBack(this); });
        orderMenu->addAction("Bring Forward", this, [mw, this]() { mw->bringForward(this); });
        orderMenu->addAction("Send Backward", this, [mw, this]() { mw->sendBackward(this); });
    }

    // Group / Ungroup
    if (mw) {
        menu.addSeparator();
        if (mw->selectedOverlays().size() >= 2)
            menu.addAction("Group", mw, &MainWindow::groupSelectedOverlays);
        if (mw->findGroupContaining(this))
            menu.addAction("Ungroup", mw, &MainWindow::ungroupSelectedOverlays);
    }

    menu.addSeparator();
    menu.addAction("Delete Shape", this, [this]() { emit deleteRequested(this); });

    menu.exec(event->globalPos());
}

void ShapeWidget::keyPressEvent(QKeyEvent* event) {
    if (m_selected && (event->key() == Qt::Key_Delete || event->key() == Qt::Key_Backspace)) {
        emit deleteRequested(this);
        event->accept();
        return;
    }
    QWidget::keyPressEvent(event);
}
