#ifndef VIRTUALSCROLLBAR_H
#define VIRTUALSCROLLBAR_H

#include <QWidget>
#include <QMouseEvent>
#include <QPainter>
#include <QTimer>
#include <cmath>

// ============================================================================
// VirtualScrollBar — Proportional scrollbar for millions of rows
// ============================================================================
// Standard QScrollBar can't handle 20M discrete positions smoothly.
// This widget uses a continuous 0.0–1.0 position with proportional thumb.
//
class VirtualScrollBar : public QWidget {
    Q_OBJECT

public:
    explicit VirtualScrollBar(QWidget* parent = nullptr)
        : QWidget(parent) {
        setFixedWidth(14);
        setCursor(Qt::ArrowCursor);
        setMouseTracking(true);

        // Smooth scroll timer for arrow/page jumps
        m_scrollTimer.setInterval(16); // ~60fps
        connect(&m_scrollTimer, &QTimer::timeout, this, &VirtualScrollBar::onScrollTimer);
    }

    void setTotalRows(int total) {
        m_totalRows = std::max(1, total);
        update();
    }

    void setVisibleRows(int visible) {
        m_visibleRows = std::max(1, visible);
        update();
    }

    void setPosition(double pos) {
        m_position = std::clamp(pos, 0.0, 1.0);
        update();
    }

    double position() const { return m_position; }

    int logicalRow() const {
        int maxStart = std::max(0, m_totalRows - m_visibleRows);
        return static_cast<int>(m_position * maxStart);
    }

    void scrollToRow(int row) {
        int maxStart = std::max(1, m_totalRows - m_visibleRows);
        setPosition(static_cast<double>(row) / maxStart);
        emit scrollPositionChanged(logicalRow());
    }

    void scrollByRows(int delta) {
        int maxStart = std::max(1, m_totalRows - m_visibleRows);
        int newRow = std::clamp(logicalRow() + delta, 0, maxStart);
        setPosition(static_cast<double>(newRow) / maxStart);
        emit scrollPositionChanged(logicalRow());
    }

signals:
    void scrollPositionChanged(int logicalRow);

protected:
    void paintEvent(QPaintEvent*) override {
        QPainter p(this);
        p.setRenderHint(QPainter::Antialiasing);

        int w = width();
        int h = height();
        int trackTop = ARROW_H;
        int trackBottom = h - ARROW_H;
        int trackH = trackBottom - trackTop;
        if (trackH <= 0) return;

        // Track background
        p.fillRect(0, trackTop, w, trackH, QColor(240, 240, 240));

        // Thumb
        double thumbRatio = std::max(0.02, static_cast<double>(m_visibleRows) / m_totalRows);
        int thumbH = std::max(MIN_THUMB_H, static_cast<int>(trackH * thumbRatio));
        int scrollableH = trackH - thumbH;
        int thumbY = trackTop + static_cast<int>(m_position * scrollableH);

        QColor thumbColor = m_dragging ? QColor(100, 100, 100) : (m_hoverThumb ? QColor(140, 140, 140) : QColor(180, 180, 180));
        p.setPen(Qt::NoPen);
        p.setBrush(thumbColor);
        p.drawRoundedRect(2, thumbY, w - 4, thumbH, 3, 3);

        // Up/down arrows
        p.setBrush(QColor(180, 180, 180));
        // Up arrow
        QPolygon upArrow;
        upArrow << QPoint(w / 2, 3) << QPoint(w - 3, ARROW_H - 3) << QPoint(3, ARROW_H - 3);
        p.drawPolygon(upArrow);
        // Down arrow
        QPolygon downArrow;
        downArrow << QPoint(w / 2, h - 3) << QPoint(w - 3, h - ARROW_H + 3) << QPoint(3, h - ARROW_H + 3);
        p.drawPolygon(downArrow);
    }

    void mousePressEvent(QMouseEvent* e) override {
        if (e->button() != Qt::LeftButton) return;

        int y = e->pos().y();
        int h = height();
        int trackTop = ARROW_H;
        int trackBottom = h - ARROW_H;

        // Click on up arrow
        if (y < ARROW_H) {
            scrollByRows(-3);
            m_scrollDirection = -3;
            m_scrollTimer.start();
            return;
        }
        // Click on down arrow
        if (y > h - ARROW_H) {
            scrollByRows(3);
            m_scrollDirection = 3;
            m_scrollTimer.start();
            return;
        }

        // Check if click is on thumb
        int trackH = trackBottom - trackTop;
        double thumbRatio = std::max(0.02, static_cast<double>(m_visibleRows) / m_totalRows);
        int thumbH = std::max(MIN_THUMB_H, static_cast<int>(trackH * thumbRatio));
        int scrollableH = trackH - thumbH;
        int thumbY = trackTop + static_cast<int>(m_position * scrollableH);

        if (y >= thumbY && y <= thumbY + thumbH) {
            // Start thumb drag
            m_dragging = true;
            m_dragOffset = y - thumbY;
            update();
        } else if (y < thumbY) {
            // Page up
            scrollByRows(-m_visibleRows);
            m_scrollDirection = -m_visibleRows;
            m_scrollTimer.start();
        } else {
            // Page down
            scrollByRows(m_visibleRows);
            m_scrollDirection = m_visibleRows;
            m_scrollTimer.start();
        }
    }

    void mouseMoveEvent(QMouseEvent* e) override {
        if (m_dragging) {
            int y = e->pos().y();
            int trackTop = ARROW_H;
            int trackBottom = height() - ARROW_H;
            int trackH = trackBottom - trackTop;
            double thumbRatio = std::max(0.02, static_cast<double>(m_visibleRows) / m_totalRows);
            int thumbH = std::max(MIN_THUMB_H, static_cast<int>(trackH * thumbRatio));
            int scrollableH = trackH - thumbH;

            if (scrollableH > 0) {
                double newPos = static_cast<double>(y - trackTop - m_dragOffset) / scrollableH;
                setPosition(newPos);
                emit scrollPositionChanged(logicalRow());
            }
        } else {
            // Hover detection for thumb highlight
            int y = e->pos().y();
            int trackTop = ARROW_H;
            int trackH = height() - 2 * ARROW_H;
            double thumbRatio = std::max(0.02, static_cast<double>(m_visibleRows) / m_totalRows);
            int thumbH = std::max(MIN_THUMB_H, static_cast<int>(trackH * thumbRatio));
            int scrollableH = trackH - thumbH;
            int thumbY = trackTop + static_cast<int>(m_position * scrollableH);
            bool onThumb = (y >= thumbY && y <= thumbY + thumbH);
            if (onThumb != m_hoverThumb) {
                m_hoverThumb = onThumb;
                update();
            }
        }
    }

    void mouseReleaseEvent(QMouseEvent*) override {
        m_dragging = false;
        m_scrollTimer.stop();
        update();
    }

    void wheelEvent(QWheelEvent* e) override {
        int delta = -e->angleDelta().y() / 40; // ~3 rows per notch
        scrollByRows(delta);
        e->accept();
    }

private:
    static constexpr int ARROW_H = 16;
    static constexpr int MIN_THUMB_H = 20;

    int m_totalRows = 1000;
    int m_visibleRows = 50;
    double m_position = 0.0;

    bool m_dragging = false;
    int m_dragOffset = 0;
    bool m_hoverThumb = false;

    int m_scrollDirection = 0;
    QTimer m_scrollTimer;

    void onScrollTimer() {
        scrollByRows(m_scrollDirection);
    }
};

#endif // VIRTUALSCROLLBAR_H
