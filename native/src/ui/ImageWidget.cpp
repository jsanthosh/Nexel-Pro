#include "ImageWidget.h"

#include <QPainter>
#include <QMouseEvent>
#include <QContextMenuEvent>
#include <QMenu>
#include <QBuffer>
#include <QFile>

// ---------------------------------------------------------------------------
// Constructor
// ---------------------------------------------------------------------------

ImageWidget::ImageWidget(QWidget* parent)
    : QWidget(parent)
{
    setMinimumSize(50, 50);
    resize(200, 200);
    setMouseTracking(true);
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

void ImageWidget::setImage(const QPixmap& pixmap)
{
    m_pixmap = pixmap;
    update();
}

void ImageWidget::setImageFromFile(const QString& filePath)
{
    QPixmap px(filePath);
    if (px.isNull())
        return;

    m_pixmap = px;
    m_config.filePath = filePath;

    // Store raw bytes as PNG so we can round-trip via config
    QByteArray bytes;
    QBuffer buffer(&bytes);
    buffer.open(QIODevice::WriteOnly);
    m_pixmap.save(&buffer, "PNG");
    m_config.imageData = bytes;

    // Resize widget to image size, capped at 600x600
    QSize sz = m_pixmap.size();
    if (sz.width() > 600 || sz.height() > 600)
        sz.scale(600, 600, Qt::KeepAspectRatio);
    resize(sz);

    update();
}

void ImageWidget::setConfig(const ImageConfig& config)
{
    m_config = config;
    if (!m_config.imageData.isEmpty()) {
        m_pixmap.loadFromData(m_config.imageData);
    }
    update();
}

void ImageWidget::setSelected(bool selected)
{
    m_selected = selected;
    if (!m_selected)
        setCursor(Qt::ArrowCursor);
    update();
}

// ---------------------------------------------------------------------------
// Paint
// ---------------------------------------------------------------------------

void ImageWidget::paintEvent(QPaintEvent* /*event*/)
{
    QPainter p(this);
    p.setRenderHint(QPainter::Antialiasing, false);
    p.setRenderHint(QPainter::SmoothPixmapTransform, true);

    const QRect r = rect();

    if (m_pixmap.isNull()) {
        // Gray placeholder
        p.fillRect(r, QColor(0xF2, 0xF4, 0xF7));
        p.setPen(QColor(0x66, 0x6C, 0x7E));
        QFont f = p.font();
        f.setPointSize(11);
        p.setFont(f);
        p.drawText(r, Qt::AlignCenter, QStringLiteral("No Image"));
    } else {
        // Draw the pixmap scaled and centered
        QPixmap scaled = m_pixmap.scaled(r.size(), Qt::KeepAspectRatio,
                                         Qt::SmoothTransformation);
        int x = (r.width()  - scaled.width())  / 2;
        int y = (r.height() - scaled.height()) / 2;
        p.drawPixmap(x, y, scaled);
    }

    // 1px border
    p.setPen(QPen(QColor(0xD0, 0xD5, 0xDD), 1));
    p.setBrush(Qt::NoBrush);
    p.drawRect(r.adjusted(0, 0, -1, -1));

    // Selection decoration
    if (m_selected) {
        p.setPen(QPen(QColor(0x00, 0x78, 0xD4), 2));
        p.setBrush(Qt::NoBrush);
        p.drawRect(r.adjusted(1, 1, -1, -1));
        drawSelectionHandles(p);
    }
}

// ---------------------------------------------------------------------------
// Selection handles
// ---------------------------------------------------------------------------

void ImageWidget::drawSelectionHandles(QPainter& p)
{
    const int hs = HANDLE_SIZE;
    const int half = hs / 2;
    const QRect r = rect();

    const QPoint centers[8] = {
        r.topLeft()     + QPoint(half, half),                          // TopLeft
        r.topRight()    + QPoint(-half + 1, half),                     // TopRight
        r.bottomLeft()  + QPoint(half, -half + 1),                     // BottomLeft
        r.bottomRight() + QPoint(-half + 1, -half + 1),               // BottomRight
        QPoint(r.center().x(), r.top()    + half),                     // Top
        QPoint(r.center().x(), r.bottom() - half + 1),                // Bottom
        QPoint(r.left()  + half,     r.center().y()),                  // Left
        QPoint(r.right() - half + 1, r.center().y()),                  // Right
    };

    p.setPen(QPen(QColor(0x00, 0x78, 0xD4), 1));
    p.setBrush(Qt::white);

    for (const auto& c : centers) {
        p.drawRect(c.x() - half, c.y() - half, hs, hs);
    }
}

ImageWidget::ResizeHandle ImageWidget::hitTestHandle(const QPoint& pos) const
{
    if (!m_selected)
        return None;

    const int hs = HANDLE_SIZE;
    const int half = hs / 2;
    const QRect r = rect();

    struct { QPoint center; ResizeHandle handle; } handles[8] = {
        { r.topLeft()     + QPoint(half, half),                    TopLeft     },
        { r.topRight()    + QPoint(-half + 1, half),               TopRight    },
        { r.bottomLeft()  + QPoint(half, -half + 1),               BottomLeft  },
        { r.bottomRight() + QPoint(-half + 1, -half + 1),         BottomRight },
        { QPoint(r.center().x(), r.top()    + half),               Top         },
        { QPoint(r.center().x(), r.bottom() - half + 1),          Bottom      },
        { QPoint(r.left()  + half,     r.center().y()),            Left        },
        { QPoint(r.right() - half + 1, r.center().y()),           Right       },
    };

    for (const auto& h : handles) {
        QRect hitRect(h.center.x() - half - 2, h.center.y() - half - 2,
                      hs + 4, hs + 4);
        if (hitRect.contains(pos))
            return h.handle;
    }
    return None;
}

void ImageWidget::updateCursorForHandle(ResizeHandle handle)
{
    switch (handle) {
    case TopLeft:     case BottomRight: setCursor(Qt::SizeFDiagCursor); break;
    case TopRight:    case BottomLeft:  setCursor(Qt::SizeBDiagCursor); break;
    case Top:         case Bottom:      setCursor(Qt::SizeVerCursor);   break;
    case Left:        case Right:       setCursor(Qt::SizeHorCursor);   break;
    case None:
    default:
        setCursor(m_selected ? Qt::SizeAllCursor : Qt::ArrowCursor);
        break;
    }
}

// ---------------------------------------------------------------------------
// Mouse handling
// ---------------------------------------------------------------------------

void ImageWidget::mousePressEvent(QMouseEvent* event)
{
    if (event->button() != Qt::LeftButton) {
        QWidget::mousePressEvent(event);
        return;
    }

    emit imageSelected(this);

    // Check resize handles first
    ResizeHandle handle = hitTestHandle(event->pos());
    if (handle != None) {
        m_resizing = true;
        m_activeHandle = handle;
        m_dragStart = event->globalPosition().toPoint();
        m_resizeStartGeometry = geometry();
        return;
    }

    // Otherwise start drag
    m_dragging = true;
    m_dragOffset = event->pos();
}

void ImageWidget::mouseMoveEvent(QMouseEvent* event)
{
    if (m_resizing) {
        QPoint delta = event->globalPosition().toPoint() - m_dragStart;
        QRect geo = m_resizeStartGeometry;

        switch (m_activeHandle) {
        case TopLeft:
            geo.setTopLeft(geo.topLeft() + delta);
            break;
        case TopRight:
            geo.setTopRight(geo.topRight() + delta);
            break;
        case BottomLeft:
            geo.setBottomLeft(geo.bottomLeft() + delta);
            break;
        case BottomRight:
            geo.setBottomRight(geo.bottomRight() + delta);
            break;
        case Top:
            geo.setTop(geo.top() + delta.y());
            break;
        case Bottom:
            geo.setBottom(geo.bottom() + delta.y());
            break;
        case Left:
            geo.setLeft(geo.left() + delta.x());
            break;
        case Right:
            geo.setRight(geo.right() + delta.x());
            break;
        case None:
        default:
            break;
        }

        // Enforce minimum size
        if (geo.width() < 50) {
            if (m_activeHandle == TopLeft || m_activeHandle == BottomLeft || m_activeHandle == Left)
                geo.setLeft(geo.right() - 50);
            else
                geo.setRight(geo.left() + 50);
        }
        if (geo.height() < 50) {
            if (m_activeHandle == TopLeft || m_activeHandle == TopRight || m_activeHandle == Top)
                geo.setTop(geo.bottom() - 50);
            else
                geo.setBottom(geo.top() + 50);
        }

        setGeometry(geo);
        return;
    }

    if (m_dragging) {
        QPoint newPos = mapToParent(event->pos()) - m_dragOffset;
        move(newPos);
        emit imageMoved(this);
        return;
    }

    // Hover: update cursor based on handle under pointer
    updateCursorForHandle(hitTestHandle(event->pos()));
}

void ImageWidget::mouseReleaseEvent(QMouseEvent* event)
{
    if (event->button() == Qt::LeftButton) {
        if (m_dragging)
            emit imageMoved(this);
        m_dragging = false;
        m_resizing = false;
        m_activeHandle = None;
    }
    QWidget::mouseReleaseEvent(event);
}

void ImageWidget::mouseDoubleClickEvent(QMouseEvent* /*event*/)
{
    emit editRequested(this);
}

// ---------------------------------------------------------------------------
// Context menu
// ---------------------------------------------------------------------------

void ImageWidget::contextMenuEvent(QContextMenuEvent* event)
{
    QMenu menu(this);

    QAction* changeAction = menu.addAction(QStringLiteral("Change Image..."));
    connect(changeAction, &QAction::triggered, this, [this]() {
        emit editRequested(this);
    });

    QAction* deleteAction = menu.addAction(QStringLiteral("Delete Image"));
    connect(deleteAction, &QAction::triggered, this, [this]() {
        emit deleteRequested(this);
    });

    menu.exec(event->globalPos());
}
