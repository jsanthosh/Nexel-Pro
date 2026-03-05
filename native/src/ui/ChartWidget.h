#ifndef CHARTWIDGET_H
#define CHARTWIDGET_H

#include <QWidget>
#include <QPoint>
#include <QRect>
#include <QString>
#include <QColor>
#include <QVector>
#include <QVariantAnimation>
#include <memory>

class Spreadsheet;
struct DocumentTheme;

// Chart types supported
enum class ChartType {
    Line,
    Bar,
    Scatter,
    Pie,
    Area,
    Donut,
    Column,
    Histogram
};

// Chart rendering backend
enum class ChartBackend {
    QtPainter,   // Default: custom QPainter rendering
    Data2App     // macOS-only: Metal+Skia via Data2App ChartView
};

// A single data series for chart rendering
struct ChartSeries {
    QString name;
    QVector<double> xValues;
    QVector<double> yValues;
    QColor color;
};

// Chart configuration
struct ChartConfig {
    ChartType type = ChartType::Column;
    QString title;
    QString xAxisTitle;
    QString yAxisTitle;
    QString dataRange;     // e.g. "A1:D10"
    bool showLegend = true;
    bool showGridLines = true;
    int themeIndex = 0;    // 0=Document Theme, 1=Excel, 2=Material, 3=Solarized, 4=Dark, 5=Mono, 6=Pastel
    QVector<ChartSeries> series;

    // Title formatting
    bool titleBold = true;
    bool titleItalic = false;
    QColor titleColor = QColor("#333333");
    QColor backgroundColor = Qt::white;

    // Series visibility (empty = all visible)
    QVector<bool> seriesVisible;

    // X-axis category labels from data range (for Data2App JSON conversion)
    QStringList categoryLabels;
};

class ChartWidget : public QWidget {
    Q_OBJECT

public:
    explicit ChartWidget(QWidget* parent = nullptr);
    virtual ~ChartWidget() = default;

    virtual void setConfig(const ChartConfig& config);
    ChartConfig config() const { return m_config; }

    void setSpreadsheet(std::shared_ptr<Spreadsheet> sheet);
    virtual void loadDataFromRange(const QString& range);
    virtual void refreshData();
    virtual void startEntryAnimation();
    virtual void updateOverlayPosition() {}  // For native backend overlay repositioning

    // Lazy loading: defer data loading until chart is visible
    void setLazyLoad(bool enabled) { m_lazyLoad = enabled; }
    bool hasLazyPending() const { return !m_pendingDataRange.isEmpty(); }
    void loadPendingData();

    // Auto-generate chart titles from data range headers (only fills empty fields)
    static void autoGenerateTitles(ChartConfig& config, std::shared_ptr<Spreadsheet> sheet);

    // Get theme colors for a given theme index
    static QVector<QColor> themeColors(int themeIndex);

    // Selection state
    bool isSelected() const { return m_selected; }
    void setSelected(bool selected);

    // Series visibility
    bool isSeriesVisible(int index) const;
    void toggleSeriesVisibility(int index);

signals:
    void chartSelected(ChartWidget* chart);
    void chartMoved(ChartWidget* chart);
    void chartResized(ChartWidget* chart);
    void editRequested(ChartWidget* chart);
    void deleteRequested(ChartWidget* chart);
    void propertiesRequested(ChartWidget* chart);

protected:
    void paintEvent(QPaintEvent* event) override;
    void mousePressEvent(QMouseEvent* event) override;
    void mouseMoveEvent(QMouseEvent* event) override;
    void mouseReleaseEvent(QMouseEvent* event) override;
    void mouseDoubleClickEvent(QMouseEvent* event) override;
    void contextMenuEvent(QContextMenuEvent* event) override;
    void keyPressEvent(QKeyEvent* event) override;

protected:
    // Accessible to subclasses (e.g. NativeChartWidget)
    void drawSelectionHandles(QPainter& p);

    ChartConfig m_config;
    std::shared_ptr<Spreadsheet> m_spreadsheet;
    bool m_selected = false;
    bool m_lazyLoad = false;
    QString m_pendingDataRange;

    static constexpr int HANDLE_SIZE = 8;
    static constexpr int TITLE_HEIGHT = 30;
    static constexpr int LEGEND_HEIGHT = 25;
    static constexpr int AXIS_MARGIN = 50;

private:
    // Rendering methods
    void drawChartBackground(QPainter& p, const QRect& area);
    void drawTitle(QPainter& p, const QRect& area);
    void drawAxes(QPainter& p, const QRect& plotArea);
    void drawGridLines(QPainter& p, const QRect& plotArea);
    void drawLegend(QPainter& p, const QRect& area);

    // Chart-type specific rendering
    void drawLineChart(QPainter& p, const QRect& plotArea);
    void drawBarChart(QPainter& p, const QRect& plotArea);
    void drawColumnChart(QPainter& p, const QRect& plotArea);
    void drawScatterChart(QPainter& p, const QRect& plotArea);
    void drawPieChart(QPainter& p, const QRect& plotArea);
    void drawAreaChart(QPainter& p, const QRect& plotArea);
    void drawDonutChart(QPainter& p, const QRect& plotArea);

    // Data helpers
    void computeAxisRange(double& minVal, double& maxVal, double& step) const;
    QRect computePlotArea() const;
    QVector<QColor> getThemeColors() const;

    // Interaction
    enum ResizeHandle { None, TopLeft, TopRight, BottomLeft, BottomRight, Top, Bottom, Left, Right };
    ResizeHandle hitTestHandle(const QPoint& pos) const;
    void updateCursorForHandle(ResizeHandle handle);

    // Legend hit testing
    struct LegendItem { QRect rect; int seriesIndex; };
    QVector<LegendItem> m_legendItems;
    int legendHitTest(const QPoint& pos) const;

    // Entry animation
    QVariantAnimation* m_entryAnim = nullptr;
    double m_animProgress = 1.0;

    // Interaction state
    bool m_dragging = false;
    bool m_resizing = false;
    ResizeHandle m_activeHandle = None;
    QPoint m_dragStart;
    QPoint m_dragOffset;
    QRect m_resizeStartGeometry;
};

#endif // CHARTWIDGET_H
