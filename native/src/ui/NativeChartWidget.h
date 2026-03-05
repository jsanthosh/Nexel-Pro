#ifndef NATIVECHARTWIDGET_H
#define NATIVECHARTWIDGET_H

#include "ChartWidget.h"

#ifdef Q_OS_MACOS

#include <atomic>

class QTimer;

class NativeChartWidget : public ChartWidget {
    Q_OBJECT

public:
    explicit NativeChartWidget(QWidget* parent = nullptr);
    ~NativeChartWidget() override;

    void setConfig(const ChartConfig& config) override;
    void setSelected(bool selected) override;
    void loadDataFromRange(const QString& range) override;
    void refreshData() override;
    void startEntryAnimation() override;
    void updateOverlayPosition() override;

protected:
    void paintEvent(QPaintEvent* event) override;
    void resizeEvent(QResizeEvent* event) override;
    void moveEvent(QMoveEvent* event) override;
    void showEvent(QShowEvent* event) override;
    void hideEvent(QHideEvent* event) override;

private:
    void createNativeView();
    void destroyNativeView();
    void updateChildWindowGeometry();
    void updateNativeConfiguration();
    void pauseMetalRendering();
    void resumeMetalRendering();

    static std::string chartConfigToJson(const ChartConfig& config);
    static std::string chartTypeToString(ChartType type);
    static std::string colorToHex(const QColor& color);
    static std::string escapeJson(const std::string& s);

    void* m_nativeChartView = nullptr;  // Opaque: ChartView* (NSView subclass)
    void* m_childWindow = nullptr;      // Opaque: NSWindow* (child window hosting ChartView)
    QTimer* m_pauseTimer = nullptr;
    QTimer* m_configPushTimer = nullptr;   // Debounce config pushes to prevent double-fire
    bool m_hasNativeView = false;
    bool m_configPending = false;       // Config stored but not yet pushed (lazy creation)

    // Limit concurrent Metal views across all instances
    static constexpr int MAX_METAL_VIEWS = 50;
    static std::atomic<int> s_metalViewCount;
};

#endif // Q_OS_MACOS
#endif // NATIVECHARTWIDGET_H
