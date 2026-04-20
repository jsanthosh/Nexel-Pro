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
            mtkView.enableSetNeedsDisplay = NO;
            mtkView.paused = NO;
            break;
        }
    }
}

// ── Overrides ───────────────────────────────────────────────────────────────

void NativeChartWidget::setSelected(bool selected)
{
    ChartWidget::setSelected(selected);
    // Don't call updateChildWindowGeometry — frame change blanks Data2App.
    // Instead, toggle a selection border overlay inside the NSWindow.
    if (m_childWindow) {
        NSWindow* overlay = (NSWindow*)m_childWindow;
        // Remove old selection border if exists (identified by accessibilityIdentifier)
        NSView* contentView = overlay.contentView;
        for (NSView* sub in [contentView.subviews copy]) {
            if ([sub.accessibilityIdentifier isEqualToString:@"selectionBorder"]) {
                [sub removeFromSuperview];
            }
        }
        if (selected) {
            // Add a transparent border view on top of the chart
            NSView* borderView = [[NSView alloc] initWithFrame:contentView.bounds];
            borderView.accessibilityIdentifier = @"selectionBorder";
            borderView.autoresizingMask = NSViewWidthSizable | NSViewHeightSizable;
            borderView.wantsLayer = YES;
            borderView.layer.borderColor = [[NSColor colorWithRed:0.26 green:0.45 blue:0.77 alpha:1.0] CGColor];
            borderView.layer.borderWidth = 2.0;
            borderView.layer.cornerRadius = 2.0;
            [contentView addSubview:borderView positioned:NSWindowAbove relativeTo:nil];
            [borderView release];
        }
    }
    update(); // repaint Qt widget for any non-Metal selection visuals
}

void NativeChartWidget::setConfig(const ChartConfig& config)
{
    m_config = config;
    // Populate legend hit-test rects for mouse interaction
    computeLegendLayout();
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
    computeLegendLayout();
    // Only push to native renderer if data was actually loaded (not deferred)
    if (!m_config.series.isEmpty()) {
        updateNativeConfiguration();
    }
}

void NativeChartWidget::refreshData()
{
    // Parent loads data into m_config.series from spreadsheet
    ChartWidget::refreshData();
    computeLegendLayout();
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
    computeLegendLayout();  // Recompute legend rects for new size
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

    // Never change overlay frame on selection — ANY frame change causes
    // Data2App to blank briefly. Selection is indicated by the Qt widget's
    // paintEvent drawing a border behind the overlay (visible at edges
    // due to overlay background being the chart, not the border area).
    QRect overlayRect = visibleRect;

    // Convert visible rect to macOS screen coordinates (bottom-left origin)
    NSWindow* overlay = (NSWindow*)m_childWindow;
    NSScreen* screen = overlay.screen ?: [NSScreen mainScreen];
    CGFloat screenH = screen.frame.size.height;
    CGFloat ox = overlayRect.x();
    CGFloat oy = screenH - overlayRect.y() - overlayRect.height();

    NSRect newFrame = NSMakeRect(ox, oy, overlayRect.width(), overlayRect.height());
    NSRect oldFrame = overlay.frame;

    // Only update frame if it actually changed — frame changes cause Data2App
    // to clear the plot area and redraw, which causes bars to disappear briefly.
    bool frameChanged = (newFrame.origin.x != oldFrame.origin.x ||
                         newFrame.origin.y != oldFrame.origin.y ||
                         newFrame.size.width != oldFrame.size.width ||
                         newFrame.size.height != oldFrame.size.height);
    if (frameChanged) {
        [overlay setFrame:newFrame display:NO]; // NO = don't trigger redisplay
    }

    // Position ChartView within the overlay to show the correct portion.
    CGFloat cvX = widgetRect.x() - overlayRect.x();
    CGFloat clipBottom = (widgetRect.y() + height()) - (overlayRect.y() + overlayRect.height());
    CGFloat cvY = -clipBottom;

    ChartView* cv = (ChartView*)m_nativeChartView;
    NSRect newCvFrame = NSMakeRect(cvX, cvY, width(), height());
    NSRect oldCvFrame = cv.frame;
    if (newCvFrame.origin.x != oldCvFrame.origin.x ||
        newCvFrame.origin.y != oldCvFrame.origin.y ||
        newCvFrame.size.width != oldCvFrame.size.width ||
        newCvFrame.size.height != oldCvFrame.size.height) {
        [cv setFrame:newCvFrame];
    }

    // Only act on visibility transitions, not every scroll
    bool wasHidden = overlay.alphaValue < 1.0;

    // Ensure overlay is visible (may have been hidden by clipping or hideEvent)
    if (wasHidden && !m_configPending) {
        overlay.alphaValue = 1.0;
    }

    // Always trigger a redraw when the overlay is visible and MTKView is paused.
    if (!m_configPending) {
        for (NSView* subview in cv.subviews) {
            if ([subview isKindOfClass:[MTKView class]]) {
                MTKView* mtkView = (MTKView*)subview;
                if (mtkView.paused) {
                    mtkView.enableSetNeedsDisplay = YES;
                    [mtkView setNeedsDisplay:YES];
                    dispatch_after(dispatch_time(DISPATCH_TIME_NOW, (int64_t)(150 * NSEC_PER_MSEC)),
                                   dispatch_get_main_queue(), ^{
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

    // Only animate on the first config push (chart creation).
    // Subsequent pushes (edits, resize, property changes) should not re-animate.
    bool animate = !m_initialAnimationDone;
    std::string json = chartConfigToJson(m_config, animate);
    m_initialAnimationDone = true;

    qDebug() << "[NativeChart] updateNativeConfiguration:"
             << "categoryLabels=" << m_config.categoryLabels.size()
             << (m_config.categoryLabels.isEmpty() ? "(EMPTY - using xValues fallback)" : m_config.categoryLabels.join(", "))
             << "series=" << m_config.series.size()
             << "dataRange=" << m_config.dataRange;

    // Defer config push to allow the NSView to be properly attached.
    connect(m_configPushTimer, &QTimer::timeout, this, [this, json, animate]() {
        if (!m_nativeChartView) return;

        qDebug() << "[NativeChart] JSON sent to Data2App:\n" << QString::fromStdString(json);

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

std::string NativeChartWidget::chartConfigToJson(const ChartConfig& config, bool animate)
{
    // MINIMAL JSON STRATEGY:
    // Only send properties that the user has explicitly configured.
    // Omit everything else so Data2App lib uses its own internal defaults.
    // This lets us observe the lib's true defaults and sync our sidepanel
    // to reflect what the lib actually renders.

    std::ostringstream j;
    j << "{\n";

    bool needsComma = false;
    auto maybeComma = [&]() { if (needsComma) j << ",\n"; };

    // ── Title (only if set by user) ──
    if (!config.title.isEmpty()) {
        maybeComma();
        j << "  \"title\": { \"text\": \"" << escapeJson(config.title.toStdString()) << "\" }";
        needsComma = true;
    }

    bool isPie = (config.type == ChartType::Pie || config.type == ChartType::Donut);

    // Pie/Donut: still need xAxis data for legend category names, but no gridlines/ticks
    if (isPie) {
        maybeComma();
        j << "  \"xAxis\": [{\n";
        j << "    \"show\": false,\n";
        j << "    \"data\": [";
        if (!config.categoryLabels.isEmpty()) {
            for (int i = 0; i < config.categoryLabels.size(); ++i) {
                if (i > 0) j << ", ";
                j << "\"" << escapeJson(config.categoryLabels[i].toStdString()) << "\"";
            }
        }
        j << "]\n  }]";
        needsComma = true;
    } else {
        // ── X Axis ──
        maybeComma();
        j << "  \"xAxis\": [{\n";
        j << "    \"gridLine\": { \"show\": " << (config.showVerticalGridLines ? "true" : "false") << " }";
        if (!config.xAxisTitle.isEmpty()) {
            j << ",\n    \"title\": { \"text\": \"" << escapeJson(config.xAxisTitle.toStdString()) << "\" }";
        }
        j << ",\n    \"data\": [";
        if (!config.categoryLabels.isEmpty()) {
            for (int i = 0; i < config.categoryLabels.size(); ++i) {
                if (i > 0) j << ", ";
                j << "\"" << escapeJson(config.categoryLabels[i].toStdString()) << "\"";
            }
        } else if (!config.series.isEmpty()) {
            for (int i = 0; i < config.series[0].xValues.size(); ++i) {
                if (i > 0) j << ", ";
                j << "\"" << config.series[0].xValues[i] << "\"";
            }
        }
        j << "]\n  }]";
        needsComma = true;

        // ── Y Axis ──
        maybeComma();
        j << "  \"yAxis\": [{\n";
        j << "    \"gridLine\": { \"show\": " << (config.showHorizontalGridLines ? "true" : "false") << " }";
        if (!config.yAxisTitle.isEmpty()) {
            j << ",\n    \"title\": { \"text\": \"" << escapeJson(config.yAxisTitle.toStdString()) << "\" }";
        }
        j << "\n  }]";
        needsComma = true;
    }

    // ── Legend ──
    maybeComma();
    j << "  \"legend\": { \"show\": " << (config.showLegend ? "true" : "false") << " }";
    needsComma = true;

    // ── plotOptions: animation + dataLabels ──
    // Data2App dataLabels schema (from BarChart.json): hAlign, vAlign, inside
    bool showLabels = (config.dataLabelPosition != DataLabelPosition::None);
    std::string hAlign = "center", vAlign = "top";
    bool inside = false;
    bool isBarChart = (config.type == ChartType::Bar);
    switch (config.dataLabelPosition) {
        case DataLabelPosition::Center:
            hAlign = "center"; vAlign = "middle"; inside = true; break;
        case DataLabelPosition::InsideEnd:
            if (isBarChart) { hAlign = "right"; vAlign = "middle"; inside = true; }
            else            { hAlign = "center"; vAlign = "top"; inside = true; }
            break;
        case DataLabelPosition::OutsideEnd:
        case DataLabelPosition::Above:
            if (isBarChart) { hAlign = "right"; vAlign = "middle"; inside = false; }
            else            { hAlign = "center"; vAlign = "top"; inside = false; }
            break;
        case DataLabelPosition::Below:
            hAlign = "center"; vAlign = "bottom"; inside = false; break;
        case DataLabelPosition::Left:
            hAlign = "left"; vAlign = "middle"; inside = false; break;
        case DataLabelPosition::Right:
            hAlign = "right"; vAlign = "middle"; inside = false; break;
        default: break;
    }

    bool isPieChart = (config.type == ChartType::Pie || config.type == ChartType::Donut);

    maybeComma();
    j << "  \"plotOptions\": {\n"
      << "    \"series\": {\n"
      << "      \"animation\": { \"show\": " << (animate ? "true" : "false") << " }";

    if (isPieChart) {
        // Pie/Donut: pick the labelKey based on user's checkbox priority.
        // Data2App JSON parser likely supports single dataLabels object only
        // (even though C++ API takes a vector).
        std::string labelKey = "name";  // default: show slice name
        if (config.dataLabelShowPercentage) labelKey = "percentage";
        else if (config.dataLabelShowValue) labelKey = "value";
        else if (config.dataLabelShowCategory) labelKey = "category";

        j << ",\n      \"dataLabels\": { "
          << "\"show\": true"
          << ", \"labelKey\": \"" << labelKey << "\""
          << ", \"showConnector\": true"
          << ", \"fontSize\": 11"
          << " }";
    } else {
        // Column/bar/line/area: single dataLabels object
        j << ",\n      \"dataLabels\": { "
          << "\"show\": " << (showLabels ? "true" : "false")
          << ", \"hAlign\": \"" << hAlign << "\""
          << ", \"vAlign\": \"" << vAlign << "\""
          << ", \"inside\": " << (inside ? "true" : "false")
          << " }";
    }
    j << "\n    }\n  }";
    needsComma = true;

    // ── Series ──
    std::string typeStr = chartTypeToString(config.type);
    maybeComma();
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
        // Pie/Donut: one color per slice, show each slice in the legend.
        if (config.type == ChartType::Donut) {
            j << "      \"innerRadius\": \"50%\",\n";
            j << "      \"colorByPoint\": true,\n";
            j << "      \"showInLegend\": true,\n";
        } else if (config.type == ChartType::Pie) {
            j << "      \"innerRadius\": 0,\n";
            j << "      \"colorByPoint\": true,\n";
            j << "      \"showInLegend\": true,\n";
        }
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
    // Data2App lib supported types: area, column, line, pie, scatter.
    // Donut is pie with innerRadius set.
    switch (type) {
        case ChartType::Line:      return "line";
        case ChartType::Column:    return "column";
        case ChartType::Scatter:   return "scatter";
        case ChartType::Area:      return "area";
        case ChartType::Bar:       return "column";
        case ChartType::Pie:       return "pie";
        case ChartType::Donut:     return "pie";   // same 'pie' type, innerRadius makes donut
        case ChartType::Histogram: return "column";
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
