#include "NativeChartWidget.h"

#ifdef Q_OS_MACOS

#import "ChartView.h"
#import <MetalKit/MetalKit.h>
#import <Cocoa/Cocoa.h>
#include <QPainter>
#include <QResizeEvent>
#include <QMoveEvent>
#include <QShowEvent>
#include <QHideEvent>
#include <QTimer>
#include <sstream>
#include <cmath>

// Static: track how many live Metal views exist
std::atomic<int> NativeChartWidget::s_metalViewCount{0};

// ── Constructor / Destructor ────────────────────────────────────────────────

NativeChartWidget::NativeChartWidget(QWidget* parent)
    : ChartWidget(parent)
{
    // Pause timer: after animation completes, stop the 60fps Metal render loop
    m_pauseTimer = new QTimer(this);
    m_pauseTimer->setSingleShot(true);
    m_pauseTimer->setInterval(2000);  // 2s covers animation duration + delay
    connect(m_pauseTimer, &QTimer::timeout, this, &NativeChartWidget::pauseMetalRendering);

    // Config push timer: debounce rapid config updates (prevents double-fire)
    m_configPushTimer = new QTimer(this);
    m_configPushTimer->setSingleShot(true);
    m_configPushTimer->setInterval(200);

    // Metal view creation is deferred until the chart is visible in the viewport
    // (lazy rendering). createNativeView() is called from updateChildWindowGeometry()
    // when the chart first enters the visible area.
}

NativeChartWidget::~NativeChartWidget()
{
    destroyNativeView();
}

void NativeChartWidget::createNativeView()
{
    if (m_hasNativeView) return;

    // Create ChartView with origin (0,0) — important because ChartView.mm
    // passes its frame (including origin) to MetalView's init.
    int cw = qMax(width(), 100);
    int ch = qMax(height(), 100);
    ChartView* chartView = [[ChartView alloc] initWithFrame:NSMakeRect(0, 0, cw, ch)];
    m_nativeChartView = (void*)chartView;

    // Configure the MTKView: pause until chart is configured, disable framebufferOnly for Skia
    for (NSView* subview in chartView.subviews) {
        if ([subview isKindOfClass:[MTKView class]]) {
            MTKView* mtkView = (MTKView*)subview;
            mtkView.framebufferOnly = NO;
            mtkView.paused = YES;  // paused until config is pushed
            mtkView.clearColor = MTLClearColorMake(1.0, 1.0, 1.0, 1.0);  // white, not black
            break;
        }
    }

    // Host the ChartView in its own borderless child window.
    // This completely isolates the Metal/CAMetalLayer rendering pipeline
    // from Qt's backing store, which would otherwise paint over the view.
    NSWindow* overlay = [[NSWindow alloc]
        initWithContentRect:NSMakeRect(0, 0, cw, ch)
                  styleMask:NSWindowStyleMaskBorderless
                    backing:NSBackingStoreBuffered
                      defer:NO];
    overlay.opaque = NO;
    overlay.hasShadow = NO;
    overlay.backgroundColor = [NSColor clearColor];
    overlay.ignoresMouseEvents = YES;
    overlay.contentView = chartView;
    overlay.releasedWhenClosed = NO;
    m_childWindow = (void*)overlay;

    // Attach to the Qt parent window but keep hidden until chart is configured
    NSView* qtView = (NSView*)(void*)window()->winId();
    NSWindow* qtWindow = qtView.window;
    [qtWindow addChildWindow:overlay ordered:NSWindowAbove];
    overlay.alphaValue = 0.0;  // invisible until first config push (orderOut detaches child windows)

    m_hasNativeView = true;
    s_metalViewCount.fetch_add(1);
}

void NativeChartWidget::destroyNativeView()
{
    if (!m_hasNativeView) return;

    m_pauseTimer->stop();
    m_configPushTimer->stop();
    m_configPushTimer->disconnect();

    if (m_childWindow) {
        NSWindow* overlay = (NSWindow*)m_childWindow;
        NSWindow* parent = overlay.parentWindow;
        if (parent) [parent removeChildWindow:overlay];
        [overlay close];
        [overlay release];
        m_childWindow = nullptr;
    }

    if (m_nativeChartView) {
        ChartView* cv = (ChartView*)m_nativeChartView;
        [cv release];
        m_nativeChartView = nullptr;
    }

    m_hasNativeView = false;
    s_metalViewCount.fetch_sub(1);
}

// ── Metal Render Lifecycle ──────────────────────────────────────────────────

void NativeChartWidget::pauseMetalRendering()
{
    if (!m_nativeChartView) return;
    ChartView* cv = (ChartView*)m_nativeChartView;
    for (NSView* subview in cv.subviews) {
        if ([subview isKindOfClass:[MTKView class]]) {
            MTKView* mtkView = (MTKView*)subview;
            mtkView.paused = YES;
            mtkView.enableSetNeedsDisplay = NO;
            break;
        }
    }
}

void NativeChartWidget::resumeMetalRendering()
{
    if (!m_nativeChartView) return;
    ChartView* cv = (ChartView*)m_nativeChartView;
    for (NSView* subview in cv.subviews) {
        if ([subview isKindOfClass:[MTKView class]]) {
            MTKView* mtkView = (MTKView*)subview;
            mtkView.enableSetNeedsDisplay = NO;  // continuous rendering (60fps)
            mtkView.paused = NO;
            break;
        }
    }
}

// ── Overrides ───────────────────────────────────────────────────────────────

void NativeChartWidget::setConfig(const ChartConfig& config)
{
    m_config = config;
    // Only push to native if we already have series data;
    // otherwise wait for loadDataFromRange() to provide it.
    if (!m_config.series.isEmpty()) {
        updateNativeConfiguration();
    }
}

void NativeChartWidget::loadDataFromRange(const QString& range)
{
    // Parent loads data into m_config.series from spreadsheet (may defer under lazy load)
    ChartWidget::loadDataFromRange(range);
    // Only push to native renderer if data was actually loaded (not deferred)
    if (!m_config.series.isEmpty()) {
        updateNativeConfiguration();
    }
}

void NativeChartWidget::refreshData()
{
    // Parent loads data into m_config.series from spreadsheet
    ChartWidget::refreshData();
    // Push updated data to the native renderer
    updateNativeConfiguration();
}

void NativeChartWidget::startEntryAnimation()
{
    // Data2App has built-in animation; just push config
    updateNativeConfiguration();
}

void NativeChartWidget::updateOverlayPosition()
{
    updateChildWindowGeometry();
}

void NativeChartWidget::paintEvent(QPaintEvent* event)
{
    if (!m_hasNativeView && s_metalViewCount.load() >= MAX_METAL_VIEWS) {
        // Over the Metal view limit — fall back to Qt rendering permanently
        ChartWidget::paintEvent(event);
        return;
    }

    // The Metal NSView renders the chart. We only paint selection handles on top.
    if (m_selected) {
        QPainter p(this);
        p.setRenderHint(QPainter::Antialiasing, true);
        drawSelectionHandles(p);
    }
}

void NativeChartWidget::resizeEvent(QResizeEvent* event)
{
    QWidget::resizeEvent(event);
    updateChildWindowGeometry();
}

void NativeChartWidget::moveEvent(QMoveEvent* event)
{
    QWidget::moveEvent(event);
    updateChildWindowGeometry();
}

void NativeChartWidget::showEvent(QShowEvent* event)
{
    QWidget::showEvent(event);
    updateChildWindowGeometry();
}

void NativeChartWidget::hideEvent(QHideEvent* event)
{
    QWidget::hideEvent(event);
    if (m_childWindow) {
        NSWindow* overlay = (NSWindow*)m_childWindow;
        overlay.alphaValue = 0.0;
    }
}

void NativeChartWidget::updateChildWindowGeometry()
{
    // Get viewport bounds (parent widget) for clipping
    QWidget* viewport = parentWidget();
    if (!viewport) return;

    // Widget rect in global screen coordinates
    QPoint widgetGlobal = mapToGlobal(QPoint(0, 0));
    QRect widgetRect(widgetGlobal, size());

    // Viewport rect in global screen coordinates
    QPoint vpGlobal = viewport->mapToGlobal(QPoint(0, 0));
    QRect vpRect(vpGlobal, viewport->size());

    // Visible portion = intersection of chart rect with viewport rect
    QRect visibleRect = widgetRect.intersected(vpRect);

    if (visibleRect.isEmpty()) {
        // Chart is entirely outside viewport — hide overlay, pause rendering
        if (m_childWindow) {
            NSWindow* overlay = (NSWindow*)m_childWindow;
            overlay.alphaValue = 0.0;
        }
        if (m_hasNativeView) pauseMetalRendering();
        return;
    }

    // Lazy creation: create native view when chart first becomes visible
    if (!m_hasNativeView && s_metalViewCount.load() < MAX_METAL_VIEWS) {
        createNativeView();
        if (m_configPending && !m_config.series.isEmpty()) {
            updateNativeConfiguration();
        }
    }

    if (!m_childWindow || !m_nativeChartView) return;

    // Convert visible rect to macOS screen coordinates (bottom-left origin)
    NSWindow* overlay = (NSWindow*)m_childWindow;
    NSScreen* screen = overlay.screen ?: [NSScreen mainScreen];
    CGFloat screenH = screen.frame.size.height;
    CGFloat ox = visibleRect.x();
    CGFloat oy = screenH - visibleRect.y() - visibleRect.height();

    [overlay setFrame:NSMakeRect(ox, oy, visibleRect.width(), visibleRect.height()) display:YES];

    // Position ChartView within the overlay to show the correct portion.
    // The overlay is smaller than the chart when clipped, so we offset
    // the ChartView so the visible portion aligns with the overlay.
    CGFloat cvX = widgetRect.x() - visibleRect.x();  // left clip offset
    CGFloat clipBottom = (widgetRect.y() + height()) - (visibleRect.y() + visibleRect.height());
    CGFloat cvY = -clipBottom;  // macOS y is bottom-up: bottom clip = negative y offset

    ChartView* cv = (ChartView*)m_nativeChartView;
    [cv setFrame:NSMakeRect(cvX, cvY, width(), height())];

    // Only act on visibility transitions, not every scroll
    bool wasHidden = overlay.alphaValue < 1.0;

    // Ensure overlay is visible (may have been hidden by clipping or hideEvent)
    if (wasHidden && !m_configPending) {
        overlay.alphaValue = 1.0;

        // Trigger a single-frame redraw only when chart scrolls back into view
        // after being hidden — not on every scroll repositioning.
        for (NSView* subview in cv.subviews) {
            if ([subview isKindOfClass:[MTKView class]]) {
                MTKView* mtkView = (MTKView*)subview;
                if (mtkView.paused) {
                    mtkView.enableSetNeedsDisplay = YES;
                    [mtkView setNeedsDisplay:YES];
                    // Reset after queuing the redraw so it doesn't stay in on-demand mode
                    dispatch_async(dispatch_get_main_queue(), ^{
                        mtkView.enableSetNeedsDisplay = NO;
                    });
                }
                break;
            }
        }
    }
}

// ── JSON Conversion ─────────────────────────────────────────────────────────

void NativeChartWidget::updateNativeConfiguration()
{
    if (!m_nativeChartView) {
        // Native view not yet created (lazy). Mark config as pending —
        // it will be pushed when the view is created in updateChildWindowGeometry().
        m_configPending = true;
        return;
    }

    // Keep m_configPending true until Metal has actually rendered —
    // this prevents updateChildWindowGeometry() from showing the overlay
    // prematurely during show/resize/move events.
    m_configPending = true;

    // Cancel any previous pending config push to prevent double-fire.
    // The member timer ensures only the latest config is pushed.
    m_configPushTimer->disconnect();
    m_configPushTimer->stop();

    std::string json = chartConfigToJson(m_config);

    // Defer config push to allow the NSView to be properly attached.
    connect(m_configPushTimer, &QTimer::timeout, this, [this, json]() {
        if (!m_nativeChartView) return;

        ChartView* cv = (ChartView*)m_nativeChartView;
        [cv setConfiguration:json];

        // Resume rendering AFTER config push so the chart actually draws
        resumeMetalRendering();

        // Wait for Metal to render before revealing the overlay,
        // so the user never sees a black/blank flash.
        QTimer::singleShot(100, this, [this]() {
            if (!m_nativeChartView) return;
            m_configPending = false;  // NOW safe to show
            updateChildWindowGeometry();
        });

        m_pauseTimer->start();
    }, Qt::SingleShotConnection);
    m_configPushTimer->start();
}

std::string NativeChartWidget::chartConfigToJson(const ChartConfig& config)
{
    std::ostringstream j;

    std::string bg = colorToHex(config.backgroundColor);
    std::string titleCol = colorToHex(config.titleColor);

    j << "{\n";

    // ── Chart area ──
    j << "  \"chart\": {\n";
    j << "    \"chartArea\": {\n";
    j << "      \"backgroundColor\": \"" << bg << "\",\n";
    j << "      \"borderWidth\": 1,\n";
    j << "      \"borderColor\": \"#cccccc\",\n";
    j << "      \"spacing\": [30, 50, 30, 30]\n";
    j << "    },\n";
    j << "    \"plotArea\": {\n";
    j << "      \"plotBackgroundColor\": \"" << bg << "\",\n";
    j << "      \"plotBorderWidth\": 0\n";
    j << "    }\n";
    j << "  },\n";

    // ── Title ──
    j << "  \"title\": {\n";
    j << "    \"text\": \"" << escapeJson(config.title.toStdString()) << "\",\n";
    j << "    \"fontSize\": 18,\n";
    j << "    \"align\": \"left\",\n";
    j << "    \"color\": \"" << titleCol << "\",\n";
    j << "    \"fontWeight\": \"" << (config.titleBold ? "bold" : "normal") << "\"\n";
    j << "  },\n";

    // ── X Axis ──
    j << "  \"xAxis\": [{\n";
    j << "    \"show\": true,\n";
    j << "    \"labels\": { \"show\": true, \"fontSize\": 11 },\n";
    j << "    \"gridLine\": { \"show\": " << (config.showGridLines ? "true" : "false") << " },\n";
    j << "    \"ticks\": { \"show\": false },\n";
    if (!config.xAxisTitle.isEmpty()) {
        j << "    \"title\": { \"text\": \"" << escapeJson(config.xAxisTitle.toStdString()) << "\" },\n";
    }
    j << "    \"data\": [";
    if (!config.categoryLabels.isEmpty()) {
        for (int i = 0; i < config.categoryLabels.size(); ++i) {
            if (i > 0) j << ", ";
            j << "\"" << escapeJson(config.categoryLabels[i].toStdString()) << "\"";
        }
    } else if (!config.series.isEmpty()) {
        // Fallback: use numeric x values as labels
        for (int i = 0; i < config.series[0].xValues.size(); ++i) {
            if (i > 0) j << ", ";
            j << "\"" << config.series[0].xValues[i] << "\"";
        }
    }
    j << "]\n";
    j << "  }],\n";

    // ── Y Axis ──
    j << "  \"yAxis\": [{\n";
    if (!config.yAxisTitle.isEmpty()) {
        j << "    \"title\": { \"text\": \"" << escapeJson(config.yAxisTitle.toStdString()) << "\" },\n";
    }
    j << "    \"gridLine\": { \"show\": " << (config.showGridLines ? "true" : "false") << ", \"width\": 1 },\n";
    j << "    \"labels\": { \"show\": true, \"fontSize\": 11 }\n";
    j << "  }],\n";

    // ── Legend ──
    j << "  \"legend\": { \"show\": " << (config.showLegend ? "true" : "false") << " },\n";

    // ── Plot options ──
    j << "  \"plotOptions\": {\n";
    j << "    \"series\": {\n";
    j << "      \"animation\": { \"duration\": 500, \"delay\": 200, \"show\": true },\n";
    j << "      \"dataLabels\": { \"show\": false }\n";
    j << "    }\n";
    j << "  },\n";

    // ── Series ──
    std::string typeStr = chartTypeToString(config.type);
    j << "  \"seriesOptions\": [\n";
    bool firstSeries = true;
    for (int i = 0; i < config.series.size(); ++i) {
        // Skip hidden series
        if (i < config.seriesVisible.size() && !config.seriesVisible[i]) continue;

        const auto& s = config.series[i];
        if (!firstSeries) j << ",\n";
        firstSeries = false;

        j << "    {\n";
        j << "      \"type\": \"" << typeStr << "\",\n";
        j << "      \"name\": \"" << escapeJson(s.name.toStdString()) << "\",\n";
        j << "      \"xIndex\": 0,\n";
        j << "      \"yIndex\": 0,\n";
        j << "      \"color\": \"" << colorToHex(s.color) << "\",\n";
        j << "      \"data\": [";
        for (int k = 0; k < s.yValues.size(); ++k) {
            if (k > 0) j << ", ";
            double val = s.yValues[k];
            if (std::isnan(val) || std::isinf(val)) val = 0.0;
            j << val;
        }
        j << "]\n";
        j << "    }";
    }
    j << "\n  ]\n";

    j << "}\n";
    return j.str();
}

std::string NativeChartWidget::chartTypeToString(ChartType type)
{
    switch (type) {
        case ChartType::Line:      return "line";
        case ChartType::Bar:       return "bar";
        case ChartType::Column:    return "column";
        case ChartType::Scatter:   return "scatter";
        case ChartType::Pie:       return "pie";
        case ChartType::Area:      return "area";
        case ChartType::Donut:     return "pie";       // Data2App uses pie for donut too
        case ChartType::Histogram: return "column";     // Map histogram to column
        default:                   return "column";
    }
}

std::string NativeChartWidget::colorToHex(const QColor& color)
{
    return color.name().toStdString();
}

std::string NativeChartWidget::escapeJson(const std::string& s)
{
    std::string out;
    out.reserve(s.size());
    for (char c : s) {
        switch (c) {
            case '"':  out += "\\\""; break;
            case '\\': out += "\\\\"; break;
            case '\n': out += "\\n";  break;
            case '\r': out += "\\r";  break;
            case '\t': out += "\\t";  break;
            default:   out += c;      break;
        }
    }
    return out;
}

#endif // Q_OS_MACOS
