#ifndef IMAGEWIDGET_H
#define IMAGEWIDGET_H

#include <QWidget>
#include <QPixmap>
#include <QPoint>
#include <QRect>
#include <QByteArray>

struct ImageConfig {
    QString filePath;
    QByteArray imageData;
    bool maintainAspectRatio = true;
};

class ImageWidget : public QWidget {
    Q_OBJECT
public:
    explicit ImageWidget(QWidget* parent = nullptr);

    void setImage(const QPixmap& pixmap);
    void setImageFromFile(const QString& filePath);
    void setConfig(const ImageConfig& config);
    ImageConfig config() const { return m_config; }
    QPixmap pixmap() const { return m_pixmap; }

    bool isSelected() const { return m_selected; }
    void setSelected(bool selected);

signals:
    void imageSelected(ImageWidget* image);
    void imageMoved(ImageWidget* image);
    void editRequested(ImageWidget* image);
    void deleteRequested(ImageWidget* image);

protected:
    void paintEvent(QPaintEvent* event) override;
    void mousePressEvent(QMouseEvent* event) override;
    void mouseMoveEvent(QMouseEvent* event) override;
    void mouseReleaseEvent(QMouseEvent* event) override;
    void mouseDoubleClickEvent(QMouseEvent* event) override;
    void contextMenuEvent(QContextMenuEvent* event) override;

private:
    void drawSelectionHandles(QPainter& p);

    enum ResizeHandle { None, TopLeft, TopRight, BottomLeft, BottomRight, Top, Bottom, Left, Right };
    ResizeHandle hitTestHandle(const QPoint& pos) const;
    void updateCursorForHandle(ResizeHandle handle);

    ImageConfig m_config;
    QPixmap m_pixmap;
    bool m_selected = false;
    bool m_dragging = false;
    bool m_resizing = false;
    ResizeHandle m_activeHandle = None;
    QPoint m_dragStart;
    QPoint m_dragOffset;
    QRect m_resizeStartGeometry;

    static constexpr int HANDLE_SIZE = 8;
};

#endif // IMAGEWIDGET_H
