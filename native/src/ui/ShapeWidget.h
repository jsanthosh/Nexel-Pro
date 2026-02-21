#ifndef SHAPEWIDGET_H
#define SHAPEWIDGET_H

#include <QWidget>
#include <QColor>
#include <QPoint>
#include <QString>

enum class ShapeType {
    Rectangle,
    RoundedRect,
    Circle,
    Ellipse,
    Triangle,
    Star,
    Arrow,
    Line,
    Diamond,
    Pentagon,
    Hexagon,
    Callout
};

struct ShapeConfig {
    ShapeType type = ShapeType::Rectangle;
    QColor fillColor = QColor("#4A90D9");
    QColor strokeColor = QColor("#2C5F8A");
    int strokeWidth = 2;
    float opacity = 1.0f;
    float cornerRadius = 0;
    QString text;
    QColor textColor = Qt::white;
    int fontSize = 12;
};

class ShapeWidget : public QWidget {
    Q_OBJECT

public:
    explicit ShapeWidget(QWidget* parent = nullptr);
    ~ShapeWidget() = default;

    void setConfig(const ShapeConfig& config);
    ShapeConfig config() const { return m_config; }

    bool isSelected() const { return m_selected; }
    void setSelected(bool selected);

signals:
    void shapeSelected(ShapeWidget* shape);
    void shapeMoved(ShapeWidget* shape);
    void editRequested(ShapeWidget* shape);
    void deleteRequested(ShapeWidget* shape);

protected:
    void paintEvent(QPaintEvent* event) override;
    void mousePressEvent(QMouseEvent* event) override;
    void mouseMoveEvent(QMouseEvent* event) override;
    void mouseReleaseEvent(QMouseEvent* event) override;
    void mouseDoubleClickEvent(QMouseEvent* event) override;
    void contextMenuEvent(QContextMenuEvent* event) override;

private:
    void drawShape(QPainter& p, const QRect& area);
    void drawSelectionHandles(QPainter& p);

    // Shape drawing methods
    void drawRectangle(QPainter& p, const QRect& area);
    void drawRoundedRect(QPainter& p, const QRect& area);
    void drawCircle(QPainter& p, const QRect& area);
    void drawEllipse(QPainter& p, const QRect& area);
    void drawTriangle(QPainter& p, const QRect& area);
    void drawStar(QPainter& p, const QRect& area);
    void drawArrow(QPainter& p, const QRect& area);
    void drawDiamond(QPainter& p, const QRect& area);
    void drawPentagon(QPainter& p, const QRect& area);
    void drawHexagon(QPainter& p, const QRect& area);
    void drawCallout(QPainter& p, const QRect& area);
    void drawShapeLine(QPainter& p, const QRect& area);

    // Resize handles
    enum ResizeHandle { None, TopLeft, TopRight, BottomLeft, BottomRight, Top, Bottom, Left, Right };
    ResizeHandle hitTestHandle(const QPoint& pos) const;

    ShapeConfig m_config;
    bool m_selected = false;
    bool m_dragging = false;
    bool m_resizing = false;
    ResizeHandle m_activeHandle = None;
    QPoint m_dragStart;
    QPoint m_dragOffset;
    QRect m_resizeStartGeometry;

    static constexpr int HANDLE_SIZE = 8;
};

#endif // SHAPEWIDGET_H
