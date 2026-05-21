#include "GridCanvasView.h"
#include "../core/Spreadsheet.h"
#include "../core/Cell.h"
#include "../core/CellRange.h"

#include <QPainter>
#include <QPaintEvent>
#include <QResizeEvent>
#include <QMouseEvent>
#include <QKeyEvent>
#include <QWheelEvent>
#include <QScrollBar>
#include <QApplication>

GridCanvasView::GridCanvasView(QWidget* parent) : QAbstractScrollArea(parent) {
    setFocusPolicy(Qt::StrongFocus);
    setHorizontalScrollBarPolicy(Qt::ScrollBarAlwaysOn);
    setVerticalScrollBarPolicy(Qt::ScrollBarAlwaysOn);
    setMouseTracking(false);

    // We paint everything ourselves; no widget background should leak
    // through, especially during fast scrolls.
    setAttribute(Qt::WA_OpaquePaintEvent, true);
    viewport()->setAttribute(Qt::WA_OpaquePaintEvent, true);

    updateScrollRanges();
    connect(verticalScrollBar(),   &QScrollBar::valueChanged,
            this, [this](int) { viewport()->update(); });
    connect(horizontalScrollBar(), &QScrollBar::valueChanged,
            this, [this](int) { viewport()->update(); });
}

void GridCanvasView::setSpreadsheet(std::shared_ptr<Spreadsheet> sheet) {
    m_sheet = std::move(sheet);
    m_currentRow = 0;
    m_currentCol = 0;
    updateScrollRanges();
    viewport()->update();
}

void GridCanvasView::setCurrentCell(int row, int col) {
    row = qBound(0, row, kVirtualRowCount - 1);
    col = qBound(0, col, kVirtualColumnCount - 1);
    if (row == m_currentRow && col == m_currentCol) return;
    m_currentRow = row;
    m_currentCol = col;
    emit currentCellChanged(row, col);
    viewport()->update();
}

void GridCanvasView::updateScrollRanges() {
    const int gridW = viewport()->width()  - kRowHeaderWidth;
    const int gridH = viewport()->height() - kColHeaderHeight;
    const int virtualW = kVirtualColumnCount * kDefaultColWidth;
    const int virtualH = kVirtualRowCount    * kDefaultRowHeight;

    horizontalScrollBar()->setRange(0, qMax(0, virtualW - gridW));
    verticalScrollBar()->setRange(0, qMax(0, virtualH - gridH));
    horizontalScrollBar()->setPageStep(qMax(1, gridW));
    verticalScrollBar()->setPageStep(qMax(1, gridH));
    horizontalScrollBar()->setSingleStep(kDefaultColWidth);
    verticalScrollBar()->setSingleStep(kDefaultRowHeight);
}

QString GridCanvasView::colLetter(int col) const {
    QString out;
    int c = col;
    while (c >= 0) {
        out = QChar('A' + (c % 26)) + out;
        c = c / 26 - 1;
    }
    return out;
}

int GridCanvasView::rowAt(int y) const {
    if (y < kColHeaderHeight) return -1;
    const int absY = (y - kColHeaderHeight) + verticalScrollBar()->value();
    return absY / kDefaultRowHeight;
}

int GridCanvasView::colAt(int x) const {
    if (x < kRowHeaderWidth) return -1;
    const int absX = (x - kRowHeaderWidth) + horizontalScrollBar()->value();
    return absX / kDefaultColWidth;
}

QRect GridCanvasView::cellRect(int row, int col) const {
    const int xScroll = horizontalScrollBar()->value();
    const int yScroll = verticalScrollBar()->value();
    const int x = kRowHeaderWidth   + col * kDefaultColWidth  - xScroll;
    const int y = kColHeaderHeight  + row * kDefaultRowHeight - yScroll;
    return QRect(x, y, kDefaultColWidth, kDefaultRowHeight);
}

void GridCanvasView::resizeEvent(QResizeEvent* event) {
    QAbstractScrollArea::resizeEvent(event);
    updateScrollRanges();
}

void GridCanvasView::paintEvent(QPaintEvent*) {
    QPainter p(viewport());
    p.fillRect(viewport()->rect(), Qt::white);

    if (!m_sheet) {
        // Still paint headers so the widget looks consistent when empty.
        paintHeaders(p, 0, 0, 0, 0);
        return;
    }

    const int xScroll = horizontalScrollBar()->value();
    const int yScroll = verticalScrollBar()->value();
    const int gridW = viewport()->width()  - kRowHeaderWidth;
    const int gridH = viewport()->height() - kColHeaderHeight;

    const int firstCol = qMax(0, xScroll / kDefaultColWidth);
    const int firstRow = qMax(0, yScroll / kDefaultRowHeight);
    const int lastCol  = qMin(kVirtualColumnCount - 1,
                              (xScroll + gridW) / kDefaultColWidth);
    const int lastRow  = qMin(kVirtualRowCount - 1,
                              (yScroll + gridH) / kDefaultRowHeight);

    paintCells(p, firstRow, lastRow, firstCol, lastCol);
    paintSelection(p);
    paintHeaders(p, firstRow, lastRow, firstCol, lastCol);
}

void GridCanvasView::paintHeaders(QPainter& p, int firstRow, int lastRow,
                                   int firstCol, int lastCol) {
    const QColor headerBg("#F0F2F5");
    const QColor headerFg("#475467");
    const QColor border  ("#D0D5DD");

    // Top-left dead corner.
    p.fillRect(QRect(0, 0, kRowHeaderWidth, kColHeaderHeight), headerBg);
    p.setPen(border);
    p.drawLine(kRowHeaderWidth - 1, 0, kRowHeaderWidth - 1, kColHeaderHeight);
    p.drawLine(0, kColHeaderHeight - 1, kRowHeaderWidth, kColHeaderHeight - 1);

    // Column header strip.
    p.fillRect(QRect(kRowHeaderWidth, 0,
                     viewport()->width() - kRowHeaderWidth, kColHeaderHeight),
               headerBg);
    p.setPen(headerFg);
    for (int c = firstCol; c <= lastCol; ++c) {
        const QRect r(kRowHeaderWidth + c * kDefaultColWidth - horizontalScrollBar()->value(),
                      0, kDefaultColWidth, kColHeaderHeight);
        if (c == m_currentCol) {
            p.fillRect(r, QColor("#D6E4F5"));
        }
        p.setPen(headerFg);
        p.drawText(r, Qt::AlignCenter, colLetter(c));
        p.setPen(border);
        p.drawLine(r.right(), r.top(), r.right(), r.bottom());
    }
    p.setPen(border);
    p.drawLine(0, kColHeaderHeight - 1,
               viewport()->width(), kColHeaderHeight - 1);

    // Row header strip.
    p.fillRect(QRect(0, kColHeaderHeight, kRowHeaderWidth,
                     viewport()->height() - kColHeaderHeight),
               headerBg);
    for (int r = firstRow; r <= lastRow; ++r) {
        const QRect rr(0, kColHeaderHeight + r * kDefaultRowHeight - verticalScrollBar()->value(),
                       kRowHeaderWidth, kDefaultRowHeight);
        if (r == m_currentRow) {
            p.fillRect(rr, QColor("#D6E4F5"));
        }
        p.setPen(headerFg);
        p.drawText(rr, Qt::AlignCenter, QString::number(r + 1));
        p.setPen(border);
        p.drawLine(rr.left(), rr.bottom(), rr.right(), rr.bottom());
    }
    p.setPen(border);
    p.drawLine(kRowHeaderWidth - 1, 0,
               kRowHeaderWidth - 1, viewport()->height());
}

void GridCanvasView::paintCells(QPainter& p, int firstRow, int lastRow,
                                 int firstCol, int lastCol) {
    const QColor gridColor("#E4E7EC");
    const QColor textColor("#101828");

    // Clip to the cell grid so headers stay clean.
    p.save();
    p.setClipRect(kRowHeaderWidth, kColHeaderHeight,
                  viewport()->width()  - kRowHeaderWidth,
                  viewport()->height() - kColHeaderHeight);

    // Horizontal grid lines.
    p.setPen(gridColor);
    for (int r = firstRow; r <= lastRow + 1; ++r) {
        const int y = kColHeaderHeight + r * kDefaultRowHeight - verticalScrollBar()->value();
        p.drawLine(kRowHeaderWidth, y, viewport()->width(), y);
    }
    // Vertical grid lines.
    for (int c = firstCol; c <= lastCol + 1; ++c) {
        const int x = kRowHeaderWidth + c * kDefaultColWidth - horizontalScrollBar()->value();
        p.drawLine(x, kColHeaderHeight, x, viewport()->height());
    }

    // Cell text. We iterate row-major; for populated cells we ask the
    // spreadsheet for the displayed value.
    p.setPen(textColor);
    const QFontMetrics fm = p.fontMetrics();
    for (int r = firstRow; r <= lastRow; ++r) {
        for (int c = firstCol; c <= lastCol; ++c) {
            auto cell = m_sheet->getCellIfExists(r, c);
            if (!cell.isValid()) continue;
            QString text = cell->getValue().toString();
            if (text.isEmpty()) continue;

            const QRect r0 = cellRect(r, c).adjusted(4, 0, -4, 0);
            // Elide long text horizontally; vertical centring via Qt flags.
            const QString display = fm.elidedText(text, Qt::ElideRight, r0.width());
            const bool isNumeric =
                cell->getType() == CellType::Number
                || cell->getType() == CellType::Formula;
            p.drawText(r0, Qt::AlignVCenter |
                            (isNumeric ? Qt::AlignRight : Qt::AlignLeft),
                       display);
        }
    }
    p.restore();
}

void GridCanvasView::paintSelection(QPainter& p) {
    const QRect r = cellRect(m_currentRow, m_currentCol);
    if (!r.isValid()) return;
    // Don't paint the selection over the header area.
    if (r.right()  < kRowHeaderWidth)  return;
    if (r.bottom() < kColHeaderHeight) return;

    QRect clipped = r.intersected(
        QRect(kRowHeaderWidth, kColHeaderHeight,
              viewport()->width()  - kRowHeaderWidth,
              viewport()->height() - kColHeaderHeight));

    p.save();
    QPen border(QColor("#1A7A5A"));
    border.setWidth(2);
    p.setPen(border);
    p.setBrush(Qt::NoBrush);
    p.drawRect(clipped.adjusted(0, 0, -1, -1));
    p.restore();
}

void GridCanvasView::mousePressEvent(QMouseEvent* e) {
    if (e->button() != Qt::LeftButton) {
        QAbstractScrollArea::mousePressEvent(e);
        return;
    }
    const int r = rowAt(e->position().y());
    const int c = colAt(e->position().x());
    if (r >= 0 && c >= 0) {
        setCurrentCell(r, c);
        setFocus(Qt::MouseFocusReason);
    }
}

void GridCanvasView::keyPressEvent(QKeyEvent* e) {
    int dr = 0, dc = 0;
    switch (e->key()) {
        case Qt::Key_Up:    dr = -1; break;
        case Qt::Key_Down:  dr = +1; break;
        case Qt::Key_Left:  dc = -1; break;
        case Qt::Key_Right: dc = +1; break;
        case Qt::Key_PageUp: {
            const int rowsPerPage = qMax(1, (viewport()->height() - kColHeaderHeight) / kDefaultRowHeight - 1);
            dr = -rowsPerPage; break;
        }
        case Qt::Key_PageDown: {
            const int rowsPerPage = qMax(1, (viewport()->height() - kColHeaderHeight) / kDefaultRowHeight - 1);
            dr = +rowsPerPage; break;
        }
        case Qt::Key_Home: setCurrentCell(m_currentRow, 0); return;
        case Qt::Key_End:  if (m_sheet) setCurrentCell(m_currentRow, qMax(0, m_sheet->getMaxColumn())); return;
        default:
            QAbstractScrollArea::keyPressEvent(e);
            return;
    }
    setCurrentCell(m_currentRow + dr, m_currentCol + dc);

    // Auto-scroll if the new selection falls outside the viewport.
    const QRect r = cellRect(m_currentRow, m_currentCol);
    const QRect gridRect(kRowHeaderWidth, kColHeaderHeight,
                          viewport()->width()  - kRowHeaderWidth,
                          viewport()->height() - kColHeaderHeight);
    if (r.top()    < gridRect.top())    verticalScrollBar()->setValue(verticalScrollBar()->value() - (gridRect.top()    - r.top()));
    if (r.bottom() > gridRect.bottom()) verticalScrollBar()->setValue(verticalScrollBar()->value() + (r.bottom() - gridRect.bottom()));
    if (r.left()   < gridRect.left())   horizontalScrollBar()->setValue(horizontalScrollBar()->value() - (gridRect.left()   - r.left()));
    if (r.right()  > gridRect.right())  horizontalScrollBar()->setValue(horizontalScrollBar()->value() + (r.right()  - gridRect.right()));
}

void GridCanvasView::wheelEvent(QWheelEvent* e) {
    // Honour shift+wheel for horizontal, plain wheel for vertical.
    const int deltaY = e->angleDelta().y();
    const int deltaX = e->angleDelta().x();
    if (e->modifiers() & Qt::ShiftModifier) {
        horizontalScrollBar()->setValue(horizontalScrollBar()->value() - deltaY);
    } else {
        verticalScrollBar()->setValue(verticalScrollBar()->value() - deltaY);
        if (deltaX != 0) {
            horizontalScrollBar()->setValue(horizontalScrollBar()->value() - deltaX);
        }
    }
    e->accept();
}
