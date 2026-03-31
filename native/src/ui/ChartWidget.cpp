#include "ChartWidget.h"
#include "Theme.h"
#include "MainWindow.h"
#include "../core/Spreadsheet.h"
#include "../core/DocumentTheme.h"
#include <QPainter>
#include <QPainterPath>
#include <QMouseEvent>
#include <QKeyEvent>
#include <QMenu>
#include <QApplication>
#include <QScrollBar>
#include <QAbstractScrollArea>
#include <QtMath>
#include <algorithm>
#include <cmath>
#include <vector>

// --- Chart color palettes (index 0 is "Document Theme", resolved at runtime) ---
// Indices 1-6 are fixed palettes; index 0 falls through to document theme accents.
static const QVector<QVector<QColor>> kFixedPalettes = {
    // 1: Excel (fallback for Document Theme when no spreadsheet available)
    { QColor("#4472C4"), QColor("#ED7D31"), QColor("#A5A5A5"), QColor("#FFC000"), QColor("#5B9BD5"), QColor("#70AD47") },
    // 2: Material
    { QColor("#2196F3"), QColor("#FF5722"), QColor("#4CAF50"), QColor("#FFC107"), QColor("#9C27B0"), QColor("#00BCD4") },
    // 3: Solarized
    { QColor("#268BD2"), QColor("#DC322F"), QColor("#859900"), QColor("#B58900"), QColor("#6C71C4"), QColor("#2AA198") },
    // 4: Dark
    { QColor("#00C8FF"), QColor("#FF6384"), QColor("#36A2EB"), QColor("#FFCE56"), QColor("#9966FF"), QColor("#FF9F40") },
    // 5: Monochrome
    { QColor("#333333"), QColor("#666666"), QColor("#999999"), QColor("#BBBBBB"), QColor("#444444"), QColor("#777777") },
    // 6: Pastel
    { QColor("#A8D8EA"), QColor("#FFB7B2"), QColor("#B5EAD7"), QColor("#FFDAC1"), QColor("#C7CEEA"), QColor("#E2F0CB") },
};

ChartWidget::ChartWidget(QWidget* parent)
    : QWidget(parent) {
    setMinimumSize(200, 150);
    resize(400, 300);
    setAttribute(Qt::WA_DeleteOnClose, false);
    setMouseTracking(true);
    setFocusPolicy(Qt::ClickFocus);

    // Auto-scroll timer for dragging near viewport edges
    m_autoScrollTimer = new QTimer(this);
    m_autoScrollTimer->setInterval(16);  // ~60fps for fast, smooth scrolling
    connect(m_autoScrollTimer, &QTimer::timeout, this, &ChartWidget::onAutoScroll);

    // Entry animation
    m_entryAnim = new QVariantAnimation(this);
    m_entryAnim->setDuration(800);
    m_entryAnim->setStartValue(0.0);
    m_entryAnim->setEndValue(1.0);
    m_entryAnim->setEasingCurve(QEasingCurve::OutCubic);
    connect(m_entryAnim, &QVariantAnimation::valueChanged, this, [this](const QVariant& v) {
        m_animProgress = v.toDouble();
        update();
    });
}

void ChartWidget::setConfig(const ChartConfig& config) {
    bool typeChanged = m_config.type != config.type;
    m_config = config;
    if (typeChanged && !m_config.series.isEmpty()) {
        startEntryAnimation();
    } else {
        update();
    }
}

void ChartWidget::setSpreadsheet(std::shared_ptr<Spreadsheet> sheet) {
    m_spreadsheet = sheet;
}

void ChartWidget::setSelected(bool selected) {
    m_selected = selected;
    update();
}

bool ChartWidget::isSeriesVisible(int index) const {
    // For pie/donut, index refers to data points (slices), not series
    bool isPieType = (m_config.type == ChartType::Pie || m_config.type == ChartType::Donut);
    int count = isPieType
        ? (m_config.series.isEmpty() ? 0 : m_config.series[0].yValues.size())
        : m_config.series.size();
    if (index < 0 || index >= count) return false;
    if (m_config.seriesVisible.isEmpty()) return true;
    if (index >= m_config.seriesVisible.size()) return true;
    return m_config.seriesVisible[index];
}

void ChartWidget::toggleSeriesVisibility(int index) {
    bool isPieType = (m_config.type == ChartType::Pie || m_config.type == ChartType::Donut);
    int count = isPieType
        ? (m_config.series.isEmpty() ? 0 : m_config.series[0].yValues.size())
        : m_config.series.size();
    if (index < 0 || index >= count) return;
    if (m_config.seriesVisible.isEmpty()) {
        m_config.seriesVisible.fill(true, count);
    }
    while (m_config.seriesVisible.size() < count) {
        m_config.seriesVisible.append(true);
    }
    m_config.seriesVisible[index] = !m_config.seriesVisible[index];
    update();
}

int ChartWidget::legendHitTest(const QPoint& pos) const {
    for (const auto& item : m_legendItems) {
        if (item.rect.contains(pos)) return item.seriesIndex;
    }
    return -1;
}

QVector<QColor> ChartWidget::getThemeColors() const {
    // Index 0 = Document Theme: pull accent colors from spreadsheet
    if (m_config.themeIndex == 0 && m_spreadsheet) {
        const auto& dt = m_spreadsheet->getDocumentTheme();
        return { dt.colors[4], dt.colors[5], dt.colors[6],
                 dt.colors[7], dt.colors[8], dt.colors[9] };
    }
    return themeColors(m_config.themeIndex);
}

QVector<QColor> ChartWidget::themeColors(int themeIndex) {
    // Index 0 = Document Theme (use default Office accents as fallback for static calls)
    if (themeIndex == 0) {
        const auto& dt = defaultDocumentTheme();
        return { dt.colors[4], dt.colors[5], dt.colors[6],
                 dt.colors[7], dt.colors[8], dt.colors[9] };
    }
    int idx = qBound(0, themeIndex - 1, static_cast<int>(kFixedPalettes.size()) - 1);
    return kFixedPalettes[idx];
}

// --- Parse cell references and load data from spreadsheet ---

static int colFromLetter(const QString& col) {
    int result = 0;
    for (int i = 0; i < col.length(); ++i) {
        result = result * 26 + (col[i].toUpper().unicode() - 'A' + 1);
    }
    return result - 1;
}

static void parseCellRef(const QString& ref, int& row, int& col) {
    int i = 0;
    while (i < ref.length() && ref[i].isLetter()) i++;
    col = colFromLetter(ref.left(i));
    row = ref.mid(i).toInt() - 1;
}

void ChartWidget::startEntryAnimation() {
    m_animProgress = 0.0;
    m_entryAnim->stop();
    m_entryAnim->start();
}

void ChartWidget::loadPendingData() {
    if (m_pendingDataRange.isEmpty()) return;
    QString range = m_pendingDataRange;
    m_pendingDataRange.clear();
    m_lazyLoad = false;  // disable so loadDataFromRange proceeds; stays off since data is now loaded
    loadDataFromRange(range);
}

void ChartWidget::loadDataFromRange(const QString& range) {
    if (!m_spreadsheet || range.isEmpty()) return;

    // If lazy load is enabled, defer loading until chart is visible
    if (m_lazyLoad) {
        m_pendingDataRange = range;
        m_config.dataRange = range;
        return;
    }

    m_config.dataRange = range;
    m_config.series.clear();

    // Parse range like "A1:D10"
    QStringList parts = range.split(':');
    if (parts.size() != 2) return;

    int startRow, startCol, endRow, endCol;
    parseCellRef(parts[0].trimmed(), startRow, startCol);
    parseCellRef(parts[1].trimmed(), endRow, endCol);

    if (startRow > endRow) std::swap(startRow, endRow);
    if (startCol > endCol) std::swap(startCol, endCol);

    int numCols = endCol - startCol + 1;
    int numRows = endRow - startRow + 1;
    if (numCols < 1 || numRows < 1) return;

    auto colors = getThemeColors();

    // First column = X values (categories), remaining columns = Y series
    QVector<double> xValues;
    QStringList categories;

    for (int r = startRow + 1; r <= endRow; ++r) {
        auto val = m_spreadsheet->getCellValue(CellAddress(r, startCol));
        QString text = val.toString();
        bool ok;
        double num = text.toDouble(&ok);
        xValues.append(ok ? num : static_cast<double>(r - startRow));
        categories.append(text);
    }

    m_config.categoryLabels = categories;

    // Each remaining column becomes a series
    for (int c = startCol + 1; c <= endCol; ++c) {
        ChartSeries series;
        // Header row = series name
        auto headerVal = m_spreadsheet->getCellValue(CellAddress(startRow, c));
        series.name = headerVal.toString();
        if (series.name.isEmpty()) {
            series.name = QString("Series %1").arg(c - startCol);
        }

        series.xValues = xValues;
        series.color = colors[(c - startCol - 1) % colors.size()];

        for (int r = startRow + 1; r <= endRow; ++r) {
            auto val = m_spreadsheet->getCellValue(CellAddress(r, c));
            bool ok;
            double num = val.toString().toDouble(&ok);
            series.yValues.append(ok ? num : 0.0);
        }

        m_config.series.append(series);
    }

    startEntryAnimation();
}

void ChartWidget::refreshData() {
    if (!m_spreadsheet || m_config.dataRange.isEmpty()) return;

    // Reload data without re-animating
    QString range = m_config.dataRange;
    m_config.series.clear();

    QStringList parts = range.split(':');
    if (parts.size() != 2) return;

    int startRow, startCol, endRow, endCol;
    parseCellRef(parts[0].trimmed(), startRow, startCol);
    parseCellRef(parts[1].trimmed(), endRow, endCol);

    if (startRow > endRow) std::swap(startRow, endRow);
    if (startCol > endCol) std::swap(startCol, endCol);

    int numCols = endCol - startCol + 1;
    int numRows = endRow - startRow + 1;
    if (numCols < 1 || numRows < 1) return;

    auto colors = getThemeColors();
    QVector<double> xValues;
    QStringList categories;

    for (int r = startRow + 1; r <= endRow; ++r) {
        auto val = m_spreadsheet->getCellValue(CellAddress(r, startCol));
        QString text = val.toString();
        bool ok;
        double num = text.toDouble(&ok);
        xValues.append(ok ? num : static_cast<double>(r - startRow));
        categories.append(text);
    }

    m_config.categoryLabels = categories;

    for (int c = startCol + 1; c <= endCol; ++c) {
        ChartSeries series;
        auto headerVal = m_spreadsheet->getCellValue(CellAddress(startRow, c));
        series.name = headerVal.toString();
        if (series.name.isEmpty()) series.name = QString("Series %1").arg(c - startCol);
        series.xValues = xValues;
        series.color = colors[(c - startCol - 1) % colors.size()];

        for (int r = startRow + 1; r <= endRow; ++r) {
            auto val = m_spreadsheet->getCellValue(CellAddress(r, c));
            bool ok;
            double num = val.toString().toDouble(&ok);
            series.yValues.append(ok ? num : 0.0);
        }

        m_config.series.append(series);
    }

    update(); // No animation on refresh, just redraw
}

// --- Compute axis range ---

void ChartWidget::computeAxisRange(double& minVal, double& maxVal, double& step) const {
    minVal = 0;
    maxVal = 100;

    // For 100% stacked charts, the axis is always 0-100%
    if (m_config.percentStacked) {
        minVal = 0;
        maxVal = 1.0;  // Represents 100%
        step = 0.2;
        return;
    }

    // For stacked (non-percent) charts, compute max as the largest category sum
    if (m_config.stacked) {
        int numPoints = 0;
        for (int si = 0; si < m_config.series.size(); ++si) {
            if (isSeriesVisible(si))
                numPoints = qMax(numPoints, m_config.series[si].yValues.size());
        }

        minVal = 0;
        maxVal = 0;
        for (int i = 0; i < numPoints; ++i) {
            double categorySum = 0;
            for (int si = 0; si < m_config.series.size(); ++si) {
                if (!isSeriesVisible(si)) continue;
                if (i < m_config.series[si].yValues.size())
                    categorySum += m_config.series[si].yValues[i];
            }
            maxVal = qMax(maxVal, categorySum);
        }

        if (maxVal <= 0) { maxVal = 1.0; }

        // Nice number rounding
        double range = maxVal - minVal;
        double magnitude = std::pow(10.0, std::floor(std::log10(range)));
        double residual = range / magnitude;

        if (residual <= 1.5) step = 0.2 * magnitude;
        else if (residual <= 3.0) step = 0.5 * magnitude;
        else if (residual <= 7.0) step = magnitude;
        else step = 2.0 * magnitude;

        minVal = std::floor(minVal / step) * step;
        maxVal = std::ceil(maxVal / step) * step;
        if (minVal > 0) minVal = 0;
        return;
    }

    bool first = true;
    for (int si = 0; si < m_config.series.size(); ++si) {
        if (!isSeriesVisible(si)) continue;
        for (double v : m_config.series[si].yValues) {
            if (first) {
                minVal = maxVal = v;
                first = false;
            } else {
                minVal = qMin(minVal, v);
                maxVal = qMax(maxVal, v);
            }
        }
    }

    // Line charts: dynamic y-axis from data range (better for showing trends)
    // All other charts: always start at 0 (accurate visual comparison of values)
    bool isLineChart = (m_config.type == ChartType::Line);

    if (minVal == maxVal) {
        // All values identical — create a sensible range
        if (maxVal == 0) {
            // All zeros: show 0 to 1 range for any chart type
            minVal = 0;
            maxVal = 1.0;
        } else if (minVal >= 0 && !isLineChart) {
            // Non-line charts with non-negative data: range from 0 to value
            minVal = 0;
            maxVal = maxVal * 1.5;
        } else {
            minVal -= 1;
            maxVal += 1;
        }
    }

    if (!isLineChart && minVal > 0) {
        minVal = 0;
    }

    // Nice number rounding
    double range = maxVal - minVal;
    double magnitude = std::pow(10.0, std::floor(std::log10(range)));
    double residual = range / magnitude;

    if (residual <= 1.5) step = 0.2 * magnitude;
    else if (residual <= 3.0) step = 0.5 * magnitude;
    else if (residual <= 7.0) step = magnitude;
    else step = 2.0 * magnitude;

    minVal = std::floor(minVal / step) * step;
    maxVal = std::ceil(maxVal / step) * step;

    // Line charts: snap to 0 only if close to zero
    if (isLineChart && minVal > 0 && minVal < step * 2) minVal = 0;
    // Non-line charts: ensure 0 is always included
    if (!isLineChart && minVal > 0) minVal = 0;
}

void ChartWidget::autoGenerateTitles(ChartConfig& config, std::shared_ptr<Spreadsheet> sheet) {
    if (!sheet || config.dataRange.isEmpty()) return;

    QStringList parts = config.dataRange.split(':');
    if (parts.size() != 2) return;

    int startRow, startCol, endRow, endCol;
    parseCellRef(parts[0].trimmed(), startRow, startCol);
    parseCellRef(parts[1].trimmed(), endRow, endCol);
    if (startRow > endRow) std::swap(startRow, endRow);
    if (startCol > endCol) std::swap(startCol, endCol);

    // X-axis: first column header
    QString xHeader = sheet->getCellValue(CellAddress(startRow, startCol)).toString();

    // Data column headers
    QStringList dataHeaders;
    for (int c = startCol + 1; c <= endCol; ++c) {
        auto val = sheet->getCellValue(CellAddress(startRow, c));
        if (!val.toString().isEmpty()) dataHeaders << val.toString();
    }

    if (config.xAxisTitle.isEmpty() && !xHeader.isEmpty()) {
        config.xAxisTitle = xHeader;
    }

    if (config.yAxisTitle.isEmpty() && !dataHeaders.isEmpty()) {
        if (dataHeaders.size() == 1) config.yAxisTitle = dataHeaders[0];
        else if (dataHeaders.size() <= 3) config.yAxisTitle = dataHeaders.join(" / ");
    }

    if (config.title.isEmpty()) {
        if (!dataHeaders.isEmpty() && !xHeader.isEmpty()) {
            if (dataHeaders.size() == 1) config.title = dataHeaders[0] + " by " + xHeader;
            else config.title = dataHeaders.join(" & ") + " by " + xHeader;
        } else if (!dataHeaders.isEmpty()) {
            config.title = dataHeaders.join(" & ");
        }
    }
}

QRect ChartWidget::computePlotArea() const {
    int left = AXIS_MARGIN + 10;
    int top = TITLE_HEIGHT + 5;
    // Combo charts need extra right margin for the secondary Y-axis labels
    int rightMargin = (m_config.type == ChartType::Combo) ? AXIS_MARGIN + 10 : 15;
    int right = width() - rightMargin;
    int bottom = height() - (m_config.showLegend ? LEGEND_HEIGHT + 10 : 10) - 25;
    return QRect(left, top, right - left, bottom - top);
}

// --- Paint ---

void ChartWidget::paintEvent(QPaintEvent*) {
    QPainter p(this);
    p.setRenderHint(QPainter::Antialiasing, true);

    QRect area = rect();
    drawChartBackground(p, area);
    drawTitle(p, area);

    if (m_config.series.isEmpty()) {
        // No data placeholder
        p.setPen(QColor("#999"));
        p.setFont(QFont("Arial", 11));
        p.drawText(area, Qt::AlignCenter, "No data.\nSelect a range and insert a chart.");
        if (m_selected) drawSelectionHandles(p);
        return;
    }

    QRect plotArea = computePlotArea();

    // Draw based on chart type
    switch (m_config.type) {
        case ChartType::Line:
            drawAxes(p, plotArea);
            if (m_config.showGridLines) drawGridLines(p, plotArea);
            drawLineChart(p, plotArea);
            break;
        case ChartType::Bar:
            drawAxes(p, plotArea);
            if (m_config.showGridLines) drawGridLines(p, plotArea);
            drawBarChart(p, plotArea);
            break;
        case ChartType::Column:
            drawAxes(p, plotArea);
            if (m_config.showGridLines) drawGridLines(p, plotArea);
            drawColumnChart(p, plotArea);
            break;
        case ChartType::Scatter:
            drawAxes(p, plotArea);
            if (m_config.showGridLines) drawGridLines(p, plotArea);
            drawScatterChart(p, plotArea);
            break;
        case ChartType::Pie:
            drawPieChart(p, plotArea);
            break;
        case ChartType::Area:
            drawAxes(p, plotArea);
            if (m_config.showGridLines) drawGridLines(p, plotArea);
            drawAreaChart(p, plotArea);
            break;
        case ChartType::Donut:
            drawDonutChart(p, plotArea);
            break;
        case ChartType::Histogram:
            drawAxes(p, plotArea);
            if (m_config.showGridLines) drawGridLines(p, plotArea);
            drawColumnChart(p, plotArea);
            break;
        case ChartType::Combo:
            drawComboChart(p, plotArea);
            break;
        case ChartType::Waterfall:
            drawAxes(p, plotArea);
            if (m_config.showGridLines) drawGridLines(p, plotArea);
            drawWaterfallChart(p, plotArea);
            break;
    }

    // Draw trendlines
    drawTrendlines(p, plotArea);

    // Draw data labels
    drawDataLabels(p, plotArea);

    if (m_config.showLegend) drawLegend(p, area);
    if (m_selected) {
        p.setClipping(false);
        drawSelectionHandles(p);
    }
}

void ChartWidget::drawChartBackground(QPainter& p, const QRect& area) {
    // Drop shadow
    p.setPen(Qt::NoPen);
    p.setBrush(QColor(0, 0, 0, 15));
    p.drawRoundedRect(area.adjusted(4, 4, 4, 4), 12, 12);
    p.setBrush(QColor(0, 0, 0, 10));
    p.drawRoundedRect(area.adjusted(2, 2, 2, 2), 12, 12);

    // Background with rounded corners
    p.setBrush(m_config.backgroundColor);
    p.setPen(QPen(QColor("#D0D5DD"), 1));
    p.drawRoundedRect(area.adjusted(0, 0, -1, -1), 12, 12);

    // Clip to rounded rect so content doesn't overflow corners
    QPainterPath clipPath;
    clipPath.addRoundedRect(area.adjusted(1, 1, -2, -2), 11, 11);
    p.setClipPath(clipPath);
}

void ChartWidget::drawTitle(QPainter& p, const QRect& area) {
    if (m_config.title.isEmpty()) return;
    p.setPen(m_config.titleColor);
    QFont titleFont("Arial", 13, m_config.titleBold ? QFont::Bold : QFont::Normal);
    titleFont.setItalic(m_config.titleItalic);
    p.setFont(titleFont);
    p.drawText(QRect(area.left() + 10, 5, area.width() - 20, TITLE_HEIGHT),
               Qt::AlignCenter | Qt::AlignVCenter, m_config.title);
}

void ChartWidget::drawAxes(QPainter& p, const QRect& plotArea) {
    p.setPen(QPen(QColor("#888"), 1));
    // Y axis
    p.drawLine(plotArea.left(), plotArea.top(), plotArea.left(), plotArea.bottom());
    // X axis
    p.drawLine(plotArea.left(), plotArea.bottom(), plotArea.right(), plotArea.bottom());

    // Y axis ticks and labels
    double minVal, maxVal, step;
    computeAxisRange(minVal, maxVal, step);

    p.setFont(QFont("Arial", 8));
    p.setPen(QColor("#666"));

    for (double v = minVal; v <= maxVal + step * 0.001; v += step) {
        double frac = (v - minVal) / (maxVal - minVal);
        int y = plotArea.bottom() - static_cast<int>(frac * plotArea.height());
        if (y < plotArea.top() || y > plotArea.bottom()) continue;

        p.drawLine(plotArea.left() - 4, y, plotArea.left(), y);

        QString label;
        if (m_config.percentStacked) {
            label = QString::number(static_cast<int>(v * 100)) + "%";
        } else if (std::abs(v) >= 1000000) label = QString::number(v / 1000000.0, 'f', 1) + "M";
        else if (std::abs(v) >= 1000) label = QString::number(v / 1000.0, 'f', 1) + "K";
        else label = QString::number(v, 'f', step < 1 ? 1 : 0);

        p.drawText(QRect(plotArea.left() - AXIS_MARGIN, y - 8, AXIS_MARGIN - 6, 16),
                   Qt::AlignRight | Qt::AlignVCenter, label);
    }

    // X axis category labels — centered under each bar group
    if (!m_config.series.isEmpty() && !m_config.series[0].xValues.isEmpty()) {
        int n = m_config.series[0].xValues.size();
        int maxLabels = qMax(1, plotArea.width() / 50);
        int labelStep = qMax(1, n / maxLabels);
        double groupWidth = static_cast<double>(plotArea.width()) / n;

        for (int i = 0; i < n; i += labelStep) {
            int x = plotArea.left() + static_cast<int>(i * groupWidth + groupWidth / 2);
            p.drawLine(x, plotArea.bottom(), x, plotArea.bottom() + 4);
            QString label = (i < m_config.categoryLabels.size())
                            ? m_config.categoryLabels[i]
                            : QString::number(i + 1);
            p.drawText(QRect(x - 35, plotArea.bottom() + 5, 70, 16),
                       Qt::AlignCenter, label);
        }
    }

    // Axis titles
    if (!m_config.yAxisTitle.isEmpty()) {
        p.save();
        p.translate(12, plotArea.center().y());
        p.rotate(-90);
        p.setFont(QFont("Arial", 9));
        p.drawText(QRect(-plotArea.height() / 2, -10, plotArea.height(), 20),
                   Qt::AlignCenter, m_config.yAxisTitle);
        p.restore();
    }
    if (!m_config.xAxisTitle.isEmpty()) {
        p.setFont(QFont("Arial", 9));
        p.drawText(QRect(plotArea.left(), plotArea.bottom() + 18, plotArea.width(), 16),
                   Qt::AlignCenter, m_config.xAxisTitle);
    }
}

void ChartWidget::drawGridLines(QPainter& p, const QRect& plotArea) {
    double minVal, maxVal, step;
    computeAxisRange(minVal, maxVal, step);

    p.setPen(QPen(QColor("#E8E8E8"), 1, Qt::DotLine));
    for (double v = minVal + step; v < maxVal; v += step) {
        double frac = (v - minVal) / (maxVal - minVal);
        int y = plotArea.bottom() - static_cast<int>(frac * plotArea.height());
        if (y > plotArea.top() && y < plotArea.bottom()) {
            p.drawLine(plotArea.left() + 1, y, plotArea.right(), y);
        }
    }
}

void ChartWidget::drawTrendlines(QPainter& p, const QRect& plotArea) {
    if (m_config.trendlines.isEmpty() || m_config.series.isEmpty()) return;

    double minVal, maxVal, step;
    computeAxisRange(minVal, maxVal, step);

    for (int si = 0; si < qMin(m_config.trendlines.size(), m_config.series.size()); ++si) {
        const auto& tl = m_config.trendlines[si];
        if (tl.type == TrendlineType::None) continue;
        if (!isSeriesVisible(si)) continue;
        const auto& s = m_config.series[si];
        if (s.yValues.size() < 2) continue;

        int n = s.yValues.size();

        // Calculate trendline points
        QVector<QPointF> points;
        // Number of sample points for curve rendering
        const int numSamples = 100;
        // X range with forecast support
        double xStart = -tl.forecastBackward;
        double xEnd = (n - 1) + tl.forecastForward;

        if (tl.type == TrendlineType::Linear) {
            // y = mx + b via least squares
            double sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;
            for (int i = 0; i < n; ++i) {
                sumX += i; sumY += s.yValues[i];
                sumXY += i * s.yValues[i]; sumX2 += i * i;
            }
            double denom = n * sumX2 - sumX * sumX;
            if (std::abs(denom) < 1e-12) continue;
            double m = (n * sumXY - sumX * sumY) / denom;
            double b = (sumY - m * sumX) / n;

            for (int j = 0; j <= numSamples; ++j) {
                double x = xStart + (xEnd - xStart) * j / numSamples;
                points.append(QPointF(x, m * x + b));
            }

        } else if (tl.type == TrendlineType::Exponential) {
            // y = a * e^(bx)  =>  ln(y) = ln(a) + b*x
            // Only use positive y values for ln fit
            double sumX = 0, sumLnY = 0, sumXLnY = 0, sumX2 = 0;
            int validCount = 0;
            for (int i = 0; i < n; ++i) {
                if (s.yValues[i] <= 0) continue; // skip non-positive values
                double lnY = std::log(s.yValues[i]);
                sumX += i;
                sumLnY += lnY;
                sumXLnY += i * lnY;
                sumX2 += static_cast<double>(i) * i;
                ++validCount;
            }
            if (validCount < 2) continue;
            double denom = validCount * sumX2 - sumX * sumX;
            if (std::abs(denom) < 1e-12) continue;
            double b = (validCount * sumXLnY - sumX * sumLnY) / denom;
            double lnA = (sumLnY - b * sumX) / validCount;
            double a = std::exp(lnA);

            for (int j = 0; j <= numSamples; ++j) {
                double x = xStart + (xEnd - xStart) * j / numSamples;
                double y = a * std::exp(b * x);
                points.append(QPointF(x, y));
            }

        } else if (tl.type == TrendlineType::Logarithmic) {
            // y = a * ln(x) + b
            // Use x = i+1 (1-based) to avoid ln(0)
            double sumLnX = 0, sumY = 0, sumLnXY = 0, sumLnX2 = 0;
            for (int i = 0; i < n; ++i) {
                double lnX = std::log(static_cast<double>(i + 1));
                sumLnX += lnX;
                sumY += s.yValues[i];
                sumLnXY += lnX * s.yValues[i];
                sumLnX2 += lnX * lnX;
            }
            double denom = n * sumLnX2 - sumLnX * sumLnX;
            if (std::abs(denom) < 1e-12) continue;
            double a = (n * sumLnXY - sumLnX * sumY) / denom;
            double b = (sumY - a * sumLnX) / n;

            for (int j = 0; j <= numSamples; ++j) {
                double x = xStart + (xEnd - xStart) * j / numSamples;
                double xVal = x + 1; // 1-based
                if (xVal <= 0) continue; // can't take ln of non-positive
                double y = a * std::log(xVal) + b;
                points.append(QPointF(x, y));
            }

        } else if (tl.type == TrendlineType::Power) {
            // y = a * x^b  =>  ln(y) = ln(a) + b*ln(x)
            // Use x = i+1 (1-based), only positive y values
            double sumLnX = 0, sumLnY = 0, sumLnXLnY = 0, sumLnX2 = 0;
            int validCount = 0;
            for (int i = 0; i < n; ++i) {
                if (s.yValues[i] <= 0) continue; // skip non-positive
                double lnX = std::log(static_cast<double>(i + 1));
                double lnY = std::log(s.yValues[i]);
                sumLnX += lnX;
                sumLnY += lnY;
                sumLnXLnY += lnX * lnY;
                sumLnX2 += lnX * lnX;
                ++validCount;
            }
            if (validCount < 2) continue;
            double denom = validCount * sumLnX2 - sumLnX * sumLnX;
            if (std::abs(denom) < 1e-12) continue;
            double bCoeff = (validCount * sumLnXLnY - sumLnX * sumLnY) / denom;
            double lnA = (sumLnY - bCoeff * sumLnX) / validCount;
            double a = std::exp(lnA);

            for (int j = 0; j <= numSamples; ++j) {
                double x = xStart + (xEnd - xStart) * j / numSamples;
                double xVal = x + 1; // 1-based
                if (xVal <= 0) continue;
                double y = a * std::pow(xVal, bCoeff);
                points.append(QPointF(x, y));
            }

        } else if (tl.type == TrendlineType::Polynomial) {
            // y = a0 + a1*x + a2*x^2 + ... + ak*x^k
            // Solve via normal equations: (X^T X) coeffs = X^T y
            int order = qBound(2, tl.polynomialOrder, 6);
            int cols = order + 1;

            // Build X^T X matrix and X^T y vector
            // Use std::vector for matrix storage
            std::vector<double> XtX(cols * cols, 0.0);
            std::vector<double> XtY(cols, 0.0);

            // Pre-compute x powers sums
            for (int i = 0; i < n; ++i) {
                double xi = static_cast<double>(i);
                // Compute powers of xi up to 2*order
                std::vector<double> xpow(2 * order + 1);
                xpow[0] = 1.0;
                for (int k = 1; k <= 2 * order; ++k) {
                    xpow[k] = xpow[k - 1] * xi;
                }
                for (int r = 0; r < cols; ++r) {
                    for (int c = 0; c < cols; ++c) {
                        XtX[r * cols + c] += xpow[r + c];
                    }
                    XtY[r] += xpow[r] * s.yValues[i];
                }
            }

            // Solve via Gaussian elimination with partial pivoting
            // Augmented matrix [XtX | XtY]
            std::vector<double> aug(cols * (cols + 1));
            for (int r = 0; r < cols; ++r) {
                for (int c = 0; c < cols; ++c) {
                    aug[r * (cols + 1) + c] = XtX[r * cols + c];
                }
                aug[r * (cols + 1) + cols] = XtY[r];
            }

            // Forward elimination
            bool singular = false;
            for (int k = 0; k < cols; ++k) {
                // Partial pivoting: find max in column k
                int maxRow = k;
                double maxVal2 = std::abs(aug[k * (cols + 1) + k]);
                for (int r = k + 1; r < cols; ++r) {
                    double val = std::abs(aug[r * (cols + 1) + k]);
                    if (val > maxVal2) { maxVal2 = val; maxRow = r; }
                }
                if (maxVal2 < 1e-12) { singular = true; break; }

                // Swap rows
                if (maxRow != k) {
                    for (int c = 0; c <= cols; ++c) {
                        std::swap(aug[k * (cols + 1) + c], aug[maxRow * (cols + 1) + c]);
                    }
                }

                // Eliminate below
                for (int r = k + 1; r < cols; ++r) {
                    double factor = aug[r * (cols + 1) + k] / aug[k * (cols + 1) + k];
                    for (int c = k; c <= cols; ++c) {
                        aug[r * (cols + 1) + c] -= factor * aug[k * (cols + 1) + c];
                    }
                }
            }
            if (singular) continue;

            // Back substitution
            std::vector<double> coeffs(cols, 0.0);
            for (int r = cols - 1; r >= 0; --r) {
                double sum = aug[r * (cols + 1) + cols];
                for (int c = r + 1; c < cols; ++c) {
                    sum -= aug[r * (cols + 1) + c] * coeffs[c];
                }
                coeffs[r] = sum / aug[r * (cols + 1) + r];
            }

            // Generate curve points
            for (int j = 0; j <= numSamples; ++j) {
                double x = xStart + (xEnd - xStart) * j / numSamples;
                double y = 0;
                double xpow = 1.0;
                for (int k = 0; k < cols; ++k) {
                    y += coeffs[k] * xpow;
                    xpow *= x;
                }
                points.append(QPointF(x, y));
            }

        } else if (tl.type == TrendlineType::MovingAverage) {
            int period = qMax(2, tl.movingAveragePeriod);
            for (int i = period - 1; i < n; ++i) {
                double sum = 0;
                for (int j = i - period + 1; j <= i; ++j) sum += s.yValues[j];
                points.append(QPointF(i, sum / period));
            }
        }

        // Draw the trendline
        if (points.size() < 2) continue;
        QPen pen(tl.color, tl.lineWidth, Qt::DashLine);
        p.setPen(pen);
        p.setBrush(Qt::NoBrush);

        QPainterPath path;
        for (int i = 0; i < points.size(); ++i) {
            double xFrac = (n > 1) ? points[i].x() / (n - 1) : 0.5;
            double yFrac = (points[i].y() - minVal) / (maxVal - minVal);
            int px = plotArea.left() + static_cast<int>(xFrac * plotArea.width());
            int py = plotArea.bottom() - static_cast<int>(yFrac * plotArea.height());
            px = qBound(plotArea.left(), px, plotArea.right());
            py = qBound(plotArea.top(), py, plotArea.bottom());

            if (i == 0) path.moveTo(px, py);
            else path.lineTo(px, py);
        }
        p.drawPath(path);
    }
}

void ChartWidget::drawDataLabels(QPainter& p, const QRect& plotArea) {
    if (m_config.dataLabelPosition == DataLabelPosition::None) return;
    if (m_config.series.isEmpty()) return;

    double minVal, maxVal, step;
    computeAxisRange(minVal, maxVal, step);

    QFont labelFont("Arial", 8);
    p.setFont(labelFont);
    QFontMetrics fm(labelFont);

    int numSeries = m_config.series.size();
    int numPoints = m_config.series[0].yValues.size();
    if (numPoints == 0) return;

    bool isColumn = (m_config.type == ChartType::Column || m_config.type == ChartType::Histogram);
    bool isBar = (m_config.type == ChartType::Bar);
    bool isPie = (m_config.type == ChartType::Pie || m_config.type == ChartType::Donut);

    // Calculate total for percentage
    double total = 0;
    if (m_config.dataLabelShowPercentage) {
        for (const auto& s : m_config.series)
            for (double v : s.yValues) total += std::abs(v);
    }

    for (int si = 0; si < numSeries; ++si) {
        if (!isSeriesVisible(si)) continue;
        const auto& s = m_config.series[si];
        int n = s.yValues.size();
        p.setPen(QColor("#333333"));

        for (int i = 0; i < n; ++i) {
            // Build label text
            QStringList parts;
            if (m_config.dataLabelShowSeriesName) parts << s.name;
            if (m_config.dataLabelShowCategory && i < m_config.categoryLabels.size())
                parts << m_config.categoryLabels[i];
            if (m_config.dataLabelShowValue)
                parts << QString::number(s.yValues[i], 'f', 1);
            if (m_config.dataLabelShowPercentage && total > 0)
                parts << QString::number(s.yValues[i] / total * 100, 'f', 1) + "%";

            QString label = parts.join(", ");
            if (label.isEmpty()) continue;

            int tw = fm.horizontalAdvance(label);
            int th = fm.height();

            int lx, ly;

            if (isColumn) {
                // Match exact bar geometry from drawColumnChart
                double groupWidth = static_cast<double>(plotArea.width()) / numPoints;
                double gap = groupWidth * 0.15;

                int barX, barHeight, barTop, barBottom, barCenterX;
                if (m_config.stacked || m_config.percentStacked) {
                    double fullBarWidth = groupWidth * 0.7;
                    // Compute per-category total for percent stacked
                    double catTotal = 0;
                    if (m_config.percentStacked) {
                        for (int cs = 0; cs < numSeries; ++cs) {
                            if (!isSeriesVisible(cs)) continue;
                            if (i < m_config.series[cs].yValues.size())
                                catTotal += std::abs(m_config.series[cs].yValues[i]);
                        }
                    }
                    // Compute cumulative height up to this series
                    double cumH = 0;
                    for (int ps = 0; ps < si; ++ps) {
                        if (!isSeriesVisible(ps)) continue;
                        if (i >= m_config.series[ps].yValues.size()) continue;
                        double pv = m_config.series[ps].yValues[i];
                        if (m_config.percentStacked && catTotal > 0)
                            cumH += (std::abs(pv) / catTotal) * plotArea.height();
                        else
                            cumH += ((pv - minVal) / (maxVal - minVal)) * plotArea.height();
                    }
                    double thisH;
                    if (m_config.percentStacked && catTotal > 0)
                        thisH = (std::abs(s.yValues[i]) / catTotal) * plotArea.height();
                    else
                        thisH = ((s.yValues[i] - minVal) / (maxVal - minVal)) * plotArea.height();
                    barHeight = static_cast<int>(thisH);
                    barX = plotArea.left() + static_cast<int>(i * groupWidth + gap);
                    barTop = plotArea.bottom() - static_cast<int>(cumH) - barHeight;
                    barBottom = plotArea.bottom() - static_cast<int>(cumH);
                    barCenterX = barX + static_cast<int>(fullBarWidth) / 2;
                } else {
                    double barWidth = (groupWidth * 0.7) / numSeries;
                    double yFrac = (s.yValues[i] - minVal) / (maxVal - minVal);
                    barHeight = static_cast<int>(yFrac * plotArea.height());
                    barX = plotArea.left() + static_cast<int>(i * groupWidth + gap + si * barWidth);
                    barTop = plotArea.bottom() - barHeight;
                    barBottom = plotArea.bottom();
                    barCenterX = barX + static_cast<int>(barWidth) / 2;
                }

                lx = barCenterX - tw / 2;
                switch (m_config.dataLabelPosition) {
                    case DataLabelPosition::Above:
                    case DataLabelPosition::OutsideEnd:
                        ly = barTop - th - 3;
                        break;
                    case DataLabelPosition::Below:
                        ly = barBottom + 3;
                        break;
                    case DataLabelPosition::Center:
                        ly = barTop + (barHeight - th) / 2;
                        break;
                    case DataLabelPosition::InsideEnd:
                        ly = barTop + 3;
                        break;
                    default:
                        ly = barTop - th - 3;
                        break;
                }
            } else if (isBar) {
                // Match exact bar geometry from drawBarChart
                double groupHeight = static_cast<double>(plotArea.height()) / numPoints;
                double gap = groupHeight * 0.15;

                int barW, barY, barRight, barCenterY;
                if (m_config.stacked || m_config.percentStacked) {
                    double fullBarH = groupHeight * 0.7;
                    // Compute per-category total for percent stacked
                    double catTotal = 0;
                    if (m_config.percentStacked) {
                        for (int cs = 0; cs < numSeries; ++cs) {
                            if (!isSeriesVisible(cs)) continue;
                            if (i < m_config.series[cs].yValues.size())
                                catTotal += std::abs(m_config.series[cs].yValues[i]);
                        }
                    }
                    // Compute cumulative width up to this series
                    double cumW = 0;
                    for (int ps = 0; ps < si; ++ps) {
                        if (!isSeriesVisible(ps)) continue;
                        if (i >= m_config.series[ps].yValues.size()) continue;
                        double pv = m_config.series[ps].yValues[i];
                        if (m_config.percentStacked && catTotal > 0)
                            cumW += (std::abs(pv) / catTotal) * plotArea.width();
                        else
                            cumW += ((pv - minVal) / (maxVal - minVal)) * plotArea.width();
                    }
                    double thisW;
                    if (m_config.percentStacked && catTotal > 0)
                        thisW = (std::abs(s.yValues[i]) / catTotal) * plotArea.width();
                    else
                        thisW = ((s.yValues[i] - minVal) / (maxVal - minVal)) * plotArea.width();
                    barW = static_cast<int>(thisW);
                    barY = plotArea.top() + static_cast<int>(i * groupHeight + gap);
                    barRight = plotArea.left() + static_cast<int>(cumW) + barW;
                    barCenterY = barY + static_cast<int>(fullBarH) / 2;
                } else {
                    double bh = (groupHeight * 0.7) / numSeries;
                    double xFrac = (s.yValues[i] - minVal) / (maxVal - minVal);
                    barW = static_cast<int>(xFrac * plotArea.width());
                    barY = plotArea.top() + static_cast<int>(i * groupHeight + gap + si * bh);
                    barRight = plotArea.left() + barW;
                    barCenterY = barY + static_cast<int>(bh) / 2;
                }

                ly = barCenterY - th / 2;
                switch (m_config.dataLabelPosition) {
                    case DataLabelPosition::Above:
                    case DataLabelPosition::OutsideEnd:
                        lx = barRight + 4;
                        break;
                    case DataLabelPosition::Below:
                        lx = plotArea.left() - tw - 4;
                        break;
                    case DataLabelPosition::Center:
                        lx = plotArea.left() + (barW - tw) / 2;
                        break;
                    case DataLabelPosition::InsideEnd:
                        lx = barRight - tw - 4;
                        break;
                    default:
                        lx = barRight + 4;
                        break;
                }
            } else if (isPie) {
                // For pie: skip here, handled in drawPieChart
                continue;
            } else {
                // Line, Area, Scatter -- point-based positioning
                double xFrac = (n > 1) ? static_cast<double>(i) / (n - 1) : 0.5;
                double yFrac = (s.yValues[i] - minVal) / (maxVal - minVal);
                int px = plotArea.left() + static_cast<int>(xFrac * plotArea.width());
                int py = plotArea.bottom() - static_cast<int>(yFrac * plotArea.height());

                lx = px - tw / 2;
                switch (m_config.dataLabelPosition) {
                    case DataLabelPosition::Above:
                        ly = py - th - 4;
                        break;
                    case DataLabelPosition::Below:
                        ly = py + 6;
                        break;
                    case DataLabelPosition::Center:
                        ly = py - th / 2;
                        break;
                    default:
                        ly = py - th - 4;
                        break;
                }
            }

            // Clamp to plot area
            lx = qMax(plotArea.left(), qMin(lx, plotArea.right() - tw));
            ly = qMax(plotArea.top() - th, qMin(ly, plotArea.bottom()));

            p.drawText(lx, ly, tw, th, Qt::AlignCenter, label);
        }
    }
}

void ChartWidget::drawLegend(QPainter& p, const QRect& area) {
    if (m_config.series.isEmpty()) return;

    m_legendItems.clear();
    QFont baseFont("Arial", 9);
    p.setFont(baseFont);
    int y = area.bottom() - LEGEND_HEIGHT;
    int totalWidth = 0;

    // For pie/donut: show per-slice legend items using category labels
    bool isPieType = (m_config.type == ChartType::Pie || m_config.type == ChartType::Donut);
    if (isPieType && !m_config.series.isEmpty()) {
        const auto& s = m_config.series[0];
        auto colors = getThemeColors();
        int count = s.yValues.size();

        // Read category labels from spreadsheet column A
        QStringList sliceLabels;
        if (m_spreadsheet && !m_config.dataRange.isEmpty()) {
            QStringList parts = m_config.dataRange.split(':');
            if (parts.size() == 2) {
                int sr, sc, er, ec;
                // Inline parse — first column cells are labels
                auto parseRef = [](const QString& ref, int& row, int& col) {
                    QString r = ref.trimmed().toUpper();
                    col = 0; int i = 0;
                    while (i < r.size() && r[i].isLetter()) col = col * 26 + (r[i++].unicode() - 'A');
                    row = r.mid(i).toInt() - 1;
                };
                parseRef(parts[0], sr, sc);
                parseRef(parts[1], er, ec);
                if (sr > er) std::swap(sr, er);
                for (int r = sr + 1; r <= er && sliceLabels.size() < count; ++r) {
                    auto val = m_spreadsheet->getCellValue(CellAddress(r, sc));
                    sliceLabels.append(val.toString());
                }
            }
        }

        // Measure total width
        for (int i = 0; i < count; ++i) {
            QString label = (i < sliceLabels.size() && !sliceLabels[i].isEmpty())
                ? sliceLabels[i] : QString("Slice %1").arg(i + 1);
            totalWidth += 14 + p.fontMetrics().horizontalAdvance(label) + 16;
        }

        int x = (area.width() - totalWidth) / 2;

        for (int i = 0; i < count; ++i) {
            QString label = (i < sliceLabels.size() && !sliceLabels[i].isEmpty())
                ? sliceLabels[i] : QString("Slice %1").arg(i + 1);

            bool visible = isSeriesVisible(i);
            int itemStartX = x;

            p.setPen(Qt::NoPen);
            p.setBrush(visible ? colors[i % colors.size()] : QColor("#C0C0C0"));
            p.drawRoundedRect(x, y + 4, 10, 10, 2, 2);
            x += 14;

            QFont legendFont("Arial", 9);
            legendFont.setStrikeOut(!visible);
            p.setFont(legendFont);
            p.setPen(visible ? QColor("#555") : QColor("#AAAAAA"));
            p.drawText(x, y + 13, label);
            int nameWidth = p.fontMetrics().horizontalAdvance(label);
            x += nameWidth + 16;

            LegendItem item;
            item.rect = QRect(itemStartX - 2, y, x - itemStartX + 4, LEGEND_HEIGHT);
            item.seriesIndex = i;  // For pie/donut, this is the slice index
            m_legendItems.append(item);
        }
        p.setFont(baseFont);
        return;
    }

    // Normal series-based legend
    for (const auto& s : m_config.series) {
        totalWidth += 14 + p.fontMetrics().horizontalAdvance(s.name) + 16;
    }

    int x = (area.width() - totalWidth) / 2;

    for (int i = 0; i < m_config.series.size(); ++i) {
        const auto& s = m_config.series[i];
        bool visible = isSeriesVisible(i);
        int itemStartX = x;

        // Color swatch
        p.setPen(Qt::NoPen);
        p.setBrush(visible ? s.color : QColor("#C0C0C0"));
        p.drawRoundedRect(x, y + 4, 10, 10, 2, 2);
        x += 14;

        // Series name (strikethrough if hidden)
        QFont legendFont("Arial", 9);
        legendFont.setStrikeOut(!visible);
        p.setFont(legendFont);
        p.setPen(visible ? QColor("#555") : QColor("#AAAAAA"));
        p.drawText(x, y + 13, s.name);
        int nameWidth = p.fontMetrics().horizontalAdvance(s.name);
        x += nameWidth + 16;

        // Store bounding rect for hit testing
        LegendItem item;
        item.rect = QRect(itemStartX - 2, y, x - itemStartX + 4, LEGEND_HEIGHT);
        item.seriesIndex = i;
        m_legendItems.append(item);
    }
    p.setFont(baseFont);
}

void ChartWidget::computeLegendLayout() {
    m_legendItems.clear();
    if (!m_config.showLegend || m_config.series.isEmpty()) return;

    QFont baseFont("Arial", 9);
    QFontMetrics fm(baseFont);
    QRect area = rect();
    int y = area.bottom() - LEGEND_HEIGHT;

    bool isPieType = (m_config.type == ChartType::Pie || m_config.type == ChartType::Donut);
    if (isPieType && !m_config.series.isEmpty()) {
        const auto& s = m_config.series[0];
        int count = s.yValues.size();
        QStringList sliceLabels = m_config.categoryLabels;

        int totalWidth = 0;
        for (int i = 0; i < count; ++i) {
            QString label = (i < sliceLabels.size() && !sliceLabels[i].isEmpty())
                ? sliceLabels[i] : QString("Slice %1").arg(i + 1);
            totalWidth += 14 + fm.horizontalAdvance(label) + 16;
        }
        int x = (area.width() - totalWidth) / 2;
        for (int i = 0; i < count; ++i) {
            QString label = (i < sliceLabels.size() && !sliceLabels[i].isEmpty())
                ? sliceLabels[i] : QString("Slice %1").arg(i + 1);
            int itemStartX = x;
            x += 14 + fm.horizontalAdvance(label) + 16;
            LegendItem item;
            item.rect = QRect(itemStartX - 2, y, x - itemStartX + 4, LEGEND_HEIGHT);
            item.seriesIndex = i;
            m_legendItems.append(item);
        }
        return;
    }

    int totalWidth = 0;
    for (const auto& s : m_config.series)
        totalWidth += 14 + fm.horizontalAdvance(s.name) + 16;
    int x = (area.width() - totalWidth) / 2;
    for (int i = 0; i < m_config.series.size(); ++i) {
        const auto& s = m_config.series[i];
        int itemStartX = x;
        x += 14 + fm.horizontalAdvance(s.name) + 16;
        LegendItem item;
        item.rect = QRect(itemStartX - 2, y, x - itemStartX + 4, LEGEND_HEIGHT);
        item.seriesIndex = i;
        m_legendItems.append(item);
    }
}

void ChartWidget::drawSelectionHandles(QPainter& p) {
    QColor handleColor = ThemeManager::instance().currentTheme().selectionHandleColor;
    p.setPen(QPen(handleColor, 2));
    p.setBrush(Qt::NoBrush);
    p.drawRect(rect().adjusted(1, 1, -2, -2));

    // Corner and edge handles
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

// --- Line Chart ---

void ChartWidget::drawLineChart(QPainter& p, const QRect& plotArea) {
    if (m_config.series.isEmpty()) return;

    double minVal, maxVal, step;
    computeAxisRange(minVal, maxVal, step);

    // Clip to animate left-to-right reveal
    int clipW = static_cast<int>(plotArea.width() * m_animProgress);
    p.save();
    p.setClipRect(QRect(plotArea.left(), plotArea.top() - 10, clipW + 10, plotArea.height() + 20));

    for (int si = 0; si < m_config.series.size(); ++si) {
        if (!isSeriesVisible(si)) continue;
        const auto& s = m_config.series[si];
        if (s.yValues.isEmpty()) continue;

        QPainterPath path;
        int n = s.yValues.size();

        for (int i = 0; i < n; ++i) {
            double xFrac = (n > 1) ? static_cast<double>(i) / (n - 1) : 0.5;
            double yFrac = (s.yValues[i] - minVal) / (maxVal - minVal);

            int px = plotArea.left() + static_cast<int>(xFrac * plotArea.width());
            int py = plotArea.bottom() - static_cast<int>(yFrac * plotArea.height());

            if (i == 0) path.moveTo(px, py);
            else path.lineTo(px, py);
        }

        p.setPen(QPen(s.color, 2.5, Qt::SolidLine, Qt::RoundCap, Qt::RoundJoin));
        p.setBrush(Qt::NoBrush);
        p.drawPath(path);

        // Data points
        p.setPen(QPen(s.color.darker(120), 1.5));
        p.setBrush(Qt::white);
        for (int i = 0; i < n; ++i) {
            double xFrac = (n > 1) ? static_cast<double>(i) / (n - 1) : 0.5;
            double yFrac = (s.yValues[i] - minVal) / (maxVal - minVal);
            int px = plotArea.left() + static_cast<int>(xFrac * plotArea.width());
            int py = plotArea.bottom() - static_cast<int>(yFrac * plotArea.height());
            p.drawEllipse(QPoint(px, py), 4, 4);
        }
    }

    p.restore();
}

// --- Column Chart ---

void ChartWidget::drawColumnChart(QPainter& p, const QRect& plotArea) {
    if (m_config.series.isEmpty()) return;

    double minVal, maxVal, step;
    computeAxisRange(minVal, maxVal, step);

    int numSeries = m_config.series.size();
    int numPoints = m_config.series[0].yValues.size();
    if (numPoints == 0) return;

    double groupWidth = static_cast<double>(plotArea.width()) / numPoints;
    double gap = groupWidth * 0.15;

    if (m_config.stacked || m_config.percentStacked) {
        // Stacked / 100% stacked column chart
        double barWidth = groupWidth * 0.7;

        for (int i = 0; i < numPoints; ++i) {
            // Compute category total for 100% stacked
            double categoryTotal = 0;
            if (m_config.percentStacked) {
                for (int s = 0; s < numSeries; ++s) {
                    if (!isSeriesVisible(s)) continue;
                    if (i < m_config.series[s].yValues.size())
                        categoryTotal += std::abs(m_config.series[s].yValues[i]);
                }
            }

            double cumulativeHeight = 0;
            for (int si = 0; si < numSeries; ++si) {
                if (!isSeriesVisible(si)) continue;
                const auto& s = m_config.series[si];
                if (i >= s.yValues.size()) continue;

                double value = s.yValues[i];
                double barH;
                if (m_config.percentStacked && categoryTotal > 0) {
                    double pctFrac = std::abs(value) / categoryTotal;
                    barH = pctFrac * plotArea.height() * m_animProgress;
                } else {
                    double yFrac = (value - minVal) / (maxVal - minVal);
                    barH = yFrac * plotArea.height() * m_animProgress;
                }
                int barHeight = static_cast<int>(barH);
                if (barHeight < 1) barHeight = 1;

                int x = plotArea.left() + static_cast<int>(i * groupWidth + gap);
                int y = plotArea.bottom() - static_cast<int>(cumulativeHeight) - barHeight;

                QRect barRect(x, y, static_cast<int>(barWidth) - 1, barHeight);
                p.setPen(Qt::NoPen);
                p.setBrush(s.color);
                p.drawRoundedRect(barRect, 2, 2);

                cumulativeHeight += barHeight;
            }
        }
    } else {
        // Clustered (side-by-side) column chart
        double barWidth = (groupWidth * 0.7) / numSeries;

        for (int si = 0; si < numSeries; ++si) {
            if (!isSeriesVisible(si)) continue;
            const auto& s = m_config.series[si];
            for (int i = 0; i < qMin(numPoints, s.yValues.size()); ++i) {
                double yFrac = (s.yValues[i] - minVal) / (maxVal - minVal);
                int barHeight = static_cast<int>(yFrac * plotArea.height() * m_animProgress);

                int x = plotArea.left() + static_cast<int>(i * groupWidth + gap + si * barWidth);
                int y = plotArea.bottom() - barHeight;

                QRect barRect(x, y, static_cast<int>(barWidth) - 1, barHeight);

                p.setPen(Qt::NoPen);
                p.setBrush(s.color);
                p.drawRoundedRect(barRect, 2, 2);
            }
        }
    }
}

// --- Bar Chart (horizontal) ---

void ChartWidget::drawBarChart(QPainter& p, const QRect& plotArea) {
    if (m_config.series.isEmpty()) return;

    double minVal, maxVal, step;
    computeAxisRange(minVal, maxVal, step);

    int numSeries = m_config.series.size();
    int numPoints = m_config.series[0].yValues.size();
    if (numPoints == 0) return;

    double groupHeight = static_cast<double>(plotArea.height()) / numPoints;
    double gap = groupHeight * 0.15;

    if (m_config.stacked || m_config.percentStacked) {
        // Stacked / 100% stacked bar chart (horizontal)
        double barH = groupHeight * 0.7;

        for (int i = 0; i < numPoints; ++i) {
            // Compute category total for 100% stacked
            double categoryTotal = 0;
            if (m_config.percentStacked) {
                for (int s = 0; s < numSeries; ++s) {
                    if (!isSeriesVisible(s)) continue;
                    if (i < m_config.series[s].yValues.size())
                        categoryTotal += std::abs(m_config.series[s].yValues[i]);
                }
            }

            double cumulativeWidth = 0;
            for (int si = 0; si < numSeries; ++si) {
                if (!isSeriesVisible(si)) continue;
                const auto& s = m_config.series[si];
                if (i >= s.yValues.size()) continue;

                double value = s.yValues[i];
                double barW;
                if (m_config.percentStacked && categoryTotal > 0) {
                    double pctFrac = std::abs(value) / categoryTotal;
                    barW = pctFrac * plotArea.width() * m_animProgress;
                } else {
                    double xFrac = (value - minVal) / (maxVal - minVal);
                    barW = xFrac * plotArea.width() * m_animProgress;
                }
                int barWidth = static_cast<int>(barW);
                if (barWidth < 1) barWidth = 1;

                int y = plotArea.top() + static_cast<int>(i * groupHeight + gap);
                int x = plotArea.left() + static_cast<int>(cumulativeWidth);

                QRect barRect(x, y, barWidth, static_cast<int>(barH) - 1);
                p.setPen(Qt::NoPen);
                p.setBrush(s.color);
                p.drawRoundedRect(barRect, 2, 2);

                cumulativeWidth += barWidth;
            }
        }
    } else {
        // Clustered (side-by-side) bar chart
        double barHeight = (groupHeight * 0.7) / numSeries;

        for (int si = 0; si < numSeries; ++si) {
            if (!isSeriesVisible(si)) continue;
            const auto& s = m_config.series[si];
            for (int i = 0; i < qMin(numPoints, s.yValues.size()); ++i) {
                double xFrac = (s.yValues[i] - minVal) / (maxVal - minVal);
                int barW = static_cast<int>(xFrac * plotArea.width() * m_animProgress);

                int y = plotArea.top() + static_cast<int>(i * groupHeight + gap + si * barHeight);
                int x = plotArea.left();

                QRect barRect(x, y, barW, static_cast<int>(barHeight) - 1);

                p.setPen(Qt::NoPen);
                p.setBrush(s.color);
                p.drawRoundedRect(barRect, 2, 2);
            }
        }
    }
}

// --- Scatter Chart ---

void ChartWidget::drawScatterChart(QPainter& p, const QRect& plotArea) {
    if (m_config.series.isEmpty()) return;

    double minVal, maxVal, step;
    computeAxisRange(minVal, maxVal, step);

    // Compute X range
    double xMin = 0, xMax = 1;
    bool first = true;
    for (int si = 0; si < m_config.series.size(); ++si) {
        if (!isSeriesVisible(si)) continue;
        for (double v : m_config.series[si].xValues) {
            if (first) { xMin = xMax = v; first = false; }
            else { xMin = qMin(xMin, v); xMax = qMax(xMax, v); }
        }
    }
    if (xMin == xMax) { xMin -= 1; xMax += 1; }

    int pointRadius = qMax(1, static_cast<int>(5 * m_animProgress));

    for (int si = 0; si < m_config.series.size(); ++si) {
        if (!isSeriesVisible(si)) continue;
        const auto& s = m_config.series[si];
        QColor c = s.color;
        c.setAlphaF(m_animProgress);
        p.setPen(QPen(s.color.darker(110), 1.5));
        p.setBrush(c);

        int n = qMin(s.xValues.size(), s.yValues.size());
        for (int i = 0; i < n; ++i) {
            double xFrac = (s.xValues[i] - xMin) / (xMax - xMin);
            double yFrac = (s.yValues[i] - minVal) / (maxVal - minVal);

            int px = plotArea.left() + static_cast<int>(xFrac * plotArea.width());
            int py = plotArea.bottom() - static_cast<int>(yFrac * plotArea.height());

            p.drawEllipse(QPoint(px, py), pointRadius, pointRadius);
        }
    }
}

// --- Pie Chart ---

void ChartWidget::drawPieChart(QPainter& p, const QRect& plotArea) {
    if (m_config.series.isEmpty() || m_config.series[0].yValues.isEmpty()) return;

    const auto& s = m_config.series[0];
    double total = 0;
    for (int i = 0; i < s.yValues.size(); ++i) {
        if (isSeriesVisible(i)) total += qMax(0.0, s.yValues[i]);
    }
    if (total <= 0) return;

    auto colors = getThemeColors();
    int size = qMin(plotArea.width(), plotArea.height()) - 20;
    QRect pieRect(plotArea.center().x() - size / 2, plotArea.center().y() - size / 2, size, size);

    int startAngle = 90 * 16; // Start from top
    for (int i = 0; i < s.yValues.size(); ++i) {
        if (!isSeriesVisible(i)) continue;
        double frac = qMax(0.0, s.yValues[i]) / total;
        int spanAngle = static_cast<int>(frac * 360 * 16 * m_animProgress);

        p.setPen(QPen(Qt::white, 2));
        p.setBrush(colors[i % colors.size()]);
        p.drawPie(pieRect, startAngle, -spanAngle);

        // Label (only show when animation is mostly complete)
        if (m_animProgress > 0.7) {
            double labelAlpha = (m_animProgress - 0.7) / 0.3;
            double midAngle = (startAngle - spanAngle / 2.0) / 16.0;
            double rad = qDegreesToRadians(midAngle);
            int labelR = size / 2 + 15;
            int lx = pieRect.center().x() + static_cast<int>(labelR * std::cos(rad));
            int ly = pieRect.center().y() - static_cast<int>(labelR * std::sin(rad));

            if (frac >= 0.05) {
                QColor labelColor("#555");
                labelColor.setAlphaF(labelAlpha);
                p.setPen(labelColor);
                p.setFont(QFont("Arial", 8));
                p.drawText(QRect(lx - 20, ly - 8, 40, 16), Qt::AlignCenter,
                           QString::number(frac * 100, 'f', 1) + "%");
            }
        }

        startAngle -= spanAngle;
    }
}

// --- Area Chart ---

void ChartWidget::drawAreaChart(QPainter& p, const QRect& plotArea) {
    if (m_config.series.isEmpty()) return;

    double minVal, maxVal, step;
    computeAxisRange(minVal, maxVal, step);

    int numSeries = m_config.series.size();
    int numPoints = m_config.series[0].yValues.size();
    if (numPoints == 0) return;

    if (m_config.stacked || m_config.percentStacked) {
        // Stacked / 100% stacked area chart
        // Compute cumulative Y positions bottom-up, then draw top-down so
        // the first series visually appears on top.

        // cumulativeY[i] holds the pixel Y coordinate of the top of the
        // stack built so far at data point i. Starts at the baseline.
        QVector<double> baselineY(numPoints, static_cast<double>(plotArea.bottom()));

        // Precompute category totals for 100% stacked
        QVector<double> categoryTotals(numPoints, 0.0);
        if (m_config.percentStacked) {
            for (int i = 0; i < numPoints; ++i) {
                for (int si = 0; si < numSeries; ++si) {
                    if (!isSeriesVisible(si)) continue;
                    if (i < m_config.series[si].yValues.size())
                        categoryTotals[i] += std::abs(m_config.series[si].yValues[i]);
                }
            }
        }

        // Build cumulative top-Y arrays for each series (bottom-up order)
        QVector<QVector<double>> topY(numSeries, QVector<double>(numPoints));
        QVector<double> runningY = baselineY;

        for (int si = 0; si < numSeries; ++si) {
            const auto& s = m_config.series[si];
            for (int i = 0; i < numPoints; ++i) {
                if (!isSeriesVisible(si) || i >= s.yValues.size()) {
                    topY[si][i] = runningY[i];
                    continue;
                }
                double value = s.yValues[i];
                double h;
                if (m_config.percentStacked && categoryTotals[i] > 0) {
                    double pctFrac = std::abs(value) / categoryTotals[i];
                    h = pctFrac * plotArea.height() * m_animProgress;
                } else {
                    double yFrac = (value - minVal) / (maxVal - minVal);
                    h = yFrac * plotArea.height() * m_animProgress;
                }
                topY[si][i] = runningY[i] - h;
                runningY[i] = topY[si][i];
            }
        }

        // Draw in reverse order (last series at bottom, first on top)
        for (int si = numSeries - 1; si >= 0; --si) {
            if (!isSeriesVisible(si)) continue;
            const auto& s = m_config.series[si];
            if (s.yValues.isEmpty()) continue;

            // The bottom edge for this series is the top of the previous series
            // (or the baseline for si == 0)
            QPolygonF polygon;

            // Forward pass: top edge of this series
            for (int i = 0; i < numPoints; ++i) {
                double xFrac = (numPoints > 1) ? static_cast<double>(i) / (numPoints - 1) : 0.5;
                double px = plotArea.left() + xFrac * plotArea.width();
                polygon << QPointF(px, topY[si][i]);
            }

            // Backward pass: bottom edge (top of series si-1, or baseline)
            for (int i = numPoints - 1; i >= 0; --i) {
                double xFrac = (numPoints > 1) ? static_cast<double>(i) / (numPoints - 1) : 0.5;
                double px = plotArea.left() + xFrac * plotArea.width();
                double bottomEdge = (si > 0) ? topY[si - 1][i] : static_cast<double>(plotArea.bottom());
                // For hidden earlier series, walk back to find the actual visible bottom
                if (si > 0 && !isSeriesVisible(si - 1)) {
                    bottomEdge = static_cast<double>(plotArea.bottom());
                    for (int prev = si - 1; prev >= 0; --prev) {
                        if (isSeriesVisible(prev)) {
                            bottomEdge = topY[prev][i];
                            break;
                        }
                    }
                }
                polygon << QPointF(px, bottomEdge);
            }

            QColor fill = s.color;
            fill.setAlpha(120);
            p.setPen(Qt::NoPen);
            p.setBrush(fill);
            p.drawPolygon(polygon);

            // Line on top edge
            QPainterPath line;
            for (int i = 0; i < numPoints; ++i) {
                double xFrac = (numPoints > 1) ? static_cast<double>(i) / (numPoints - 1) : 0.5;
                QPointF pt(plotArea.left() + xFrac * plotArea.width(), topY[si][i]);
                if (i == 0) line.moveTo(pt);
                else line.lineTo(pt);
            }
            p.setPen(QPen(s.color, 2));
            p.setBrush(Qt::NoBrush);
            p.drawPath(line);
        }
    } else {
        // Non-stacked area chart (original behavior)
        for (int si = m_config.series.size() - 1; si >= 0; --si) {
            if (!isSeriesVisible(si)) continue;
            const auto& s = m_config.series[si];
            if (s.yValues.isEmpty()) continue;

            int n = s.yValues.size();
            QPolygonF polygon;

            // Bottom line
            polygon << QPointF(plotArea.left(), plotArea.bottom());

            for (int i = 0; i < n; ++i) {
                double xFrac = (n > 1) ? static_cast<double>(i) / (n - 1) : 0.5;
                double yFrac = (s.yValues[i] - minVal) / (maxVal - minVal) * m_animProgress;
                polygon << QPointF(plotArea.left() + xFrac * plotArea.width(),
                                   plotArea.bottom() - yFrac * plotArea.height());
            }

            // Close at bottom
            polygon << QPointF(plotArea.right(), plotArea.bottom());

            QColor fill = s.color;
            fill.setAlpha(80);
            p.setPen(Qt::NoPen);
            p.setBrush(fill);
            p.drawPolygon(polygon);

            // Line on top
            QPainterPath line;
            for (int i = 0; i < n; ++i) {
                double xFrac = (n > 1) ? static_cast<double>(i) / (n - 1) : 0.5;
                double yFrac = (s.yValues[i] - minVal) / (maxVal - minVal) * m_animProgress;
                QPointF pt(plotArea.left() + xFrac * plotArea.width(),
                           plotArea.bottom() - yFrac * plotArea.height());
                if (i == 0) line.moveTo(pt);
                else line.lineTo(pt);
            }
            p.setPen(QPen(s.color, 2));
            p.setBrush(Qt::NoBrush);
            p.drawPath(line);
        }
    }
}

// --- Donut Chart ---

void ChartWidget::drawDonutChart(QPainter& p, const QRect& plotArea) {
    if (m_config.series.isEmpty() || m_config.series[0].yValues.isEmpty()) return;

    const auto& s = m_config.series[0];
    double total = 0;
    for (int i = 0; i < s.yValues.size(); ++i) {
        if (isSeriesVisible(i)) total += qMax(0.0, s.yValues[i]);
    }
    if (total <= 0) return;

    auto colors = getThemeColors();
    int size = qMin(plotArea.width(), plotArea.height()) - 20;
    QRect outerRect(plotArea.center().x() - size / 2, plotArea.center().y() - size / 2, size, size);
    int innerSize = size * 55 / 100;
    QRect innerRect(plotArea.center().x() - innerSize / 2, plotArea.center().y() - innerSize / 2,
                    innerSize, innerSize);

    int startAngle = 90 * 16;
    for (int i = 0; i < s.yValues.size(); ++i) {
        if (!isSeriesVisible(i)) continue;
        double frac = qMax(0.0, s.yValues[i]) / total;
        int spanAngle = static_cast<int>(frac * 360 * 16 * m_animProgress);

        QPainterPath slice;
        slice.moveTo(outerRect.center());
        slice.arcTo(outerRect, startAngle / 16.0, -spanAngle / 16.0);
        slice.closeSubpath();

        QPainterPath hole;
        hole.addEllipse(innerRect);

        QPainterPath donutSlice = slice.subtracted(hole);

        p.setPen(QPen(Qt::white, 2));
        p.setBrush(colors[i % colors.size()]);
        p.drawPath(donutSlice);

        startAngle -= spanAngle;
    }
}

// --- Waterfall Chart ---

void ChartWidget::drawWaterfallChart(QPainter& p, const QRect& plotArea) {
    if (m_config.series.isEmpty()) return;
    const auto& s = m_config.series[0]; // waterfall uses first series only
    int n = s.yValues.size();
    if (n == 0) return;

    // Compute running totals to find correct axis range
    std::vector<double> runningTotal(n);
    double cumulative = 0;
    double totalMin = 0, totalMax = 0;
    for (int i = 0; i < n; ++i) {
        double prevCumulative = cumulative;
        cumulative += s.yValues[i];
        runningTotal[i] = cumulative;
        totalMin = std::min(totalMin, std::min(cumulative, prevCumulative));
        totalMax = std::max(totalMax, std::max(cumulative, prevCumulative));
    }

    // Use computeAxisRange as baseline, then override if running totals exceed
    double minVal, maxVal, step;
    computeAxisRange(minVal, maxVal, step);
    minVal = std::min(minVal, totalMin);
    maxVal = std::max(maxVal, totalMax);
    if (maxVal == minVal) maxVal = minVal + 1;

    // Re-compute nice step for the possibly expanded range
    double range = maxVal - minVal;
    double magnitude = std::pow(10.0, std::floor(std::log10(range)));
    double residual = range / magnitude;
    if (residual <= 1.5) step = 0.2 * magnitude;
    else if (residual <= 3.0) step = 0.5 * magnitude;
    else if (residual <= 7.0) step = magnitude;
    else step = 2.0 * magnitude;
    minVal = std::floor(minVal / step) * step;
    maxVal = std::ceil(maxVal / step) * step;

    double groupWidth = static_cast<double>(plotArea.width()) / n;
    double barWidth = groupWidth * 0.6;
    double gap = groupWidth * 0.2;

    QColor posColor("#4CAF50");   // green for positive
    QColor negColor("#F44336");   // red for negative
    QColor totalColor("#2196F3"); // blue for first/last (totals)

    cumulative = 0;
    for (int i = 0; i < n; ++i) {
        double value = s.yValues[i];
        double barStart = cumulative;
        cumulative += value;
        double barEnd = cumulative;

        // Convert to pixel coordinates
        int x = plotArea.left() + static_cast<int>(i * groupWidth + gap);
        int yStart = plotArea.bottom() - static_cast<int>((barStart - minVal) / (maxVal - minVal) * plotArea.height());
        int yEnd = plotArea.bottom() - static_cast<int>((barEnd - minVal) / (maxVal - minVal) * plotArea.height());

        int barTop = std::min(yStart, yEnd);
        int barHeight = std::abs(yEnd - yStart);
        if (barHeight < 1) barHeight = 1;

        // Apply entry animation
        barHeight = static_cast<int>(barHeight * m_animProgress);
        if (barHeight < 1) barHeight = 1;

        QRect barRect(x, barTop, static_cast<int>(barWidth), barHeight);

        // Color: first bar = total, last bar = total, positive = green, negative = red
        bool isTotal = (i == 0 || i == n - 1);
        QColor color = isTotal ? totalColor : (value >= 0 ? posColor : negColor);

        p.setPen(Qt::NoPen);
        p.setBrush(color);
        p.drawRoundedRect(barRect, 2, 2);

        // Connector line to next bar (dashed, showing running total)
        if (i < n - 1) {
            int connectorY = yEnd;
            int nextX = plotArea.left() + static_cast<int>((i + 1) * groupWidth + gap);
            p.setPen(QPen(QColor("#999999"), 1, Qt::DashLine));
            p.drawLine(x + static_cast<int>(barWidth), connectorY, nextX, connectorY);
        }
    }
}

// --- Combo Chart with Secondary Y-Axis ---

void ChartWidget::drawSecondaryYAxis(QPainter& p, const QRect& plotArea,
                                      double minVal, double maxVal, double step) {
    // Draw the right-side Y axis line
    p.setPen(QPen(QColor("#888"), 1));
    p.drawLine(plotArea.right(), plotArea.top(), plotArea.right(), plotArea.bottom());

    // Y axis ticks and labels on the right side
    p.setFont(QFont("Arial", 8));
    p.setPen(QColor("#666"));

    for (double v = minVal; v <= maxVal + step * 0.001; v += step) {
        double frac = (v - minVal) / (maxVal - minVal);
        int y = plotArea.bottom() - static_cast<int>(frac * plotArea.height());
        if (y < plotArea.top() || y > plotArea.bottom()) continue;

        p.drawLine(plotArea.right(), y, plotArea.right() + 4, y);

        QString label;
        if (std::abs(v) >= 1000000) label = QString::number(v / 1000000.0, 'f', 1) + "M";
        else if (std::abs(v) >= 1000) label = QString::number(v / 1000.0, 'f', 1) + "K";
        else label = QString::number(v, 'f', step < 1 ? 1 : 0);

        p.drawText(QRect(plotArea.right() + 6, y - 8, AXIS_MARGIN - 6, 16),
                   Qt::AlignLeft | Qt::AlignVCenter, label);
    }
}

void ChartWidget::drawComboChart(QPainter& p, const QRect& plotArea) {
    if (m_config.series.isEmpty()) return;

    int numSeries = m_config.series.size();

    // Determine which series use the secondary axis
    // If useSecondaryAxis is configured, use that; otherwise heuristic: first series = primary, rest = secondary
    QVector<bool> isSecondary(numSeries, false);
    if (m_config.useSecondaryAxis.size() >= numSeries) {
        isSecondary = m_config.useSecondaryAxis;
    } else {
        // Heuristic: first series is primary (columns), remaining are secondary (lines)
        for (int i = 1; i < numSeries; ++i) {
            isSecondary[i] = true;
        }
    }

    // Compute primary axis range (from primary series only)
    double priMin = 0, priMax = 0;
    bool priFirst = true;
    for (int si = 0; si < numSeries; ++si) {
        if (isSecondary[si] || !isSeriesVisible(si)) continue;
        for (double v : m_config.series[si].yValues) {
            if (priFirst) { priMin = priMax = v; priFirst = false; }
            else { priMin = qMin(priMin, v); priMax = qMax(priMax, v); }
        }
    }
    if (priFirst) { priMin = 0; priMax = 100; }
    if (priMin > 0) priMin = 0;
    if (priMax == priMin) priMax = priMin + 1;
    // Nice rounding for primary
    {
        double range = priMax - priMin;
        double mag = std::pow(10.0, std::floor(std::log10(range)));
        double res = range / mag;
        double priStep;
        if (res <= 1.5) priStep = 0.2 * mag;
        else if (res <= 3.0) priStep = 0.5 * mag;
        else if (res <= 7.0) priStep = mag;
        else priStep = 2.0 * mag;
        priMin = std::floor(priMin / priStep) * priStep;
        priMax = std::ceil(priMax / priStep) * priStep;
        if (priMin > 0) priMin = 0;
    }

    // Compute secondary axis range (from secondary series only)
    double secMin = 0, secMax = 0;
    bool secFirst = true;
    for (int si = 0; si < numSeries; ++si) {
        if (!isSecondary[si] || !isSeriesVisible(si)) continue;
        for (double v : m_config.series[si].yValues) {
            if (secFirst) { secMin = secMax = v; secFirst = false; }
            else { secMin = qMin(secMin, v); secMax = qMax(secMax, v); }
        }
    }
    if (secFirst) { secMin = 0; secMax = 100; }
    if (secMin > 0) secMin = 0;
    if (secMax == secMin) secMax = secMin + 1;
    double secStep;
    {
        double range = secMax - secMin;
        double mag = std::pow(10.0, std::floor(std::log10(range)));
        double res = range / mag;
        if (res <= 1.5) secStep = 0.2 * mag;
        else if (res <= 3.0) secStep = 0.5 * mag;
        else if (res <= 7.0) secStep = mag;
        else secStep = 2.0 * mag;
        secMin = std::floor(secMin / secStep) * secStep;
        secMax = std::ceil(secMax / secStep) * secStep;
        if (secMin > 0) secMin = 0;
    }

    // Draw primary axis (left) and gridlines
    drawAxes(p, plotArea);
    if (m_config.showGridLines) drawGridLines(p, plotArea);

    // Draw secondary axis (right)
    drawSecondaryYAxis(p, plotArea, secMin, secMax, secStep);

    // Draw primary series as columns
    {
        int numPrimary = 0;
        for (int si = 0; si < numSeries; ++si) {
            if (!isSecondary[si] && isSeriesVisible(si)) ++numPrimary;
        }
        if (numPrimary > 0 && !m_config.series[0].yValues.isEmpty()) {
            int numPoints = m_config.series[0].yValues.size();
            double groupWidth = static_cast<double>(plotArea.width()) / numPoints;
            double gap = groupWidth * 0.15;
            double barWidth = (groupWidth * 0.7) / numPrimary;
            int priIdx = 0;
            for (int si = 0; si < numSeries; ++si) {
                if (isSecondary[si] || !isSeriesVisible(si)) continue;
                const auto& s = m_config.series[si];
                for (int i = 0; i < qMin(numPoints, s.yValues.size()); ++i) {
                    double yFrac = (s.yValues[i] - priMin) / (priMax - priMin);
                    int barHeight = static_cast<int>(yFrac * plotArea.height() * m_animProgress);
                    int x = plotArea.left() + static_cast<int>(i * groupWidth + gap + priIdx * barWidth);
                    int y = plotArea.bottom() - barHeight;
                    QRect barRect(x, y, static_cast<int>(barWidth) - 1, barHeight);
                    p.setPen(Qt::NoPen);
                    p.setBrush(s.color);
                    p.drawRoundedRect(barRect, 2, 2);
                }
                ++priIdx;
            }
        }
    }

    // Draw secondary series as lines (scaled to secondary axis)
    {
        // Clip to animate left-to-right reveal
        int clipW = static_cast<int>(plotArea.width() * m_animProgress);
        p.save();
        p.setClipRect(QRect(plotArea.left(), plotArea.top() - 10, clipW + 10, plotArea.height() + 20));

        int numPoints = m_config.series[0].yValues.size();
        for (int si = 0; si < numSeries; ++si) {
            if (!isSecondary[si] || !isSeriesVisible(si)) continue;
            const auto& s = m_config.series[si];
            if (s.yValues.isEmpty()) continue;

            int nPts = s.yValues.size();
            QPainterPath path;
            for (int i = 0; i < nPts; ++i) {
                double xFrac = (nPts > 1) ? static_cast<double>(i) / (nPts - 1) : 0.5;
                double yFrac = (s.yValues[i] - secMin) / (secMax - secMin);
                int px = plotArea.left() + static_cast<int>(xFrac * plotArea.width());
                int py = plotArea.bottom() - static_cast<int>(yFrac * plotArea.height());
                if (i == 0) path.moveTo(px, py);
                else path.lineTo(px, py);
            }
            p.setPen(QPen(s.color, 2.5, Qt::SolidLine, Qt::RoundCap, Qt::RoundJoin));
            p.setBrush(Qt::NoBrush);
            p.drawPath(path);

            // Data point markers
            if (m_config.showMarkers) {
                p.setPen(QPen(s.color.darker(120), 1.5));
                p.setBrush(Qt::white);
                for (int i = 0; i < nPts; ++i) {
                    double xFrac = (nPts > 1) ? static_cast<double>(i) / (nPts - 1) : 0.5;
                    double yFrac = (s.yValues[i] - secMin) / (secMax - secMin);
                    int px = plotArea.left() + static_cast<int>(xFrac * plotArea.width());
                    int py = plotArea.bottom() - static_cast<int>(yFrac * plotArea.height());
                    p.drawEllipse(QPoint(px, py), 4, 4);
                }
            }
        }
        p.restore();
    }
}

// --- Mouse interaction ---

ChartWidget::ResizeHandle ChartWidget::hitTestHandle(const QPoint& pos) const {
    if (!m_selected) return None;

    int w = width(), h = height();
    int hs = HANDLE_SIZE;

    if (QRect(0 - hs / 2, 0 - hs / 2, hs, hs).contains(pos)) return TopLeft;
    if (QRect(w - hs / 2, 0 - hs / 2, hs, hs).contains(pos)) return TopRight;
    if (QRect(0 - hs / 2, h - hs / 2, hs, hs).contains(pos)) return BottomLeft;
    if (QRect(w - hs / 2, h - hs / 2, hs, hs).contains(pos)) return BottomRight;
    if (QRect(w / 2 - hs / 2, 0 - hs / 2, hs, hs).contains(pos)) return Top;
    if (QRect(w / 2 - hs / 2, h - hs / 2, hs, hs).contains(pos)) return Bottom;
    if (QRect(0 - hs / 2, h / 2 - hs / 2, hs, hs).contains(pos)) return Left;
    if (QRect(w - hs / 2, h / 2 - hs / 2, hs, hs).contains(pos)) return Right;

    return None;
}

void ChartWidget::updateCursorForHandle(ResizeHandle handle) {
    switch (handle) {
        case TopLeft: case BottomRight: setCursor(Qt::SizeFDiagCursor); break;
        case TopRight: case BottomLeft: setCursor(Qt::SizeBDiagCursor); break;
        case Top: case Bottom: setCursor(Qt::SizeVerCursor); break;
        case Left: case Right: setCursor(Qt::SizeHorCursor); break;
        default: setCursor(Qt::ArrowCursor); break;
    }
}

void ChartWidget::mousePressEvent(QMouseEvent* event) {
    if (event->button() == Qt::LeftButton) {
        // Check legend click first (before drag/select)
        if (m_config.showLegend) {
            int legendIdx = legendHitTest(event->pos());
            if (legendIdx >= 0) {
                toggleSeriesVisibility(legendIdx);
                event->accept();
                return;
            }
        }

        setSelected(true);
        emit chartSelected(this);

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

void ChartWidget::mouseMoveEvent(QMouseEvent* event) {
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
            default: break;
        }

        if (geo.width() >= minimumWidth() && geo.height() >= minimumHeight()) {
            setGeometry(geo);
            emit chartResized(this);
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
            // Auto-scroll triggers when chart hits any viewport edge
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
        emit chartMoved(this);

        // Start auto-scroll as soon as chart hits viewport edge.
        // Don't stop here on !clampedAtEdge — let onAutoScroll() handle stopping
        // so a tiny mouse wiggle doesn't kill the scroll.
        if (clampedAtEdge && !m_autoScrollTimer->isActive())
            m_autoScrollTimer->start();
    } else {
        // Show pointing hand cursor over legend items
        if (m_config.showLegend && legendHitTest(event->pos()) >= 0) {
            setCursor(Qt::PointingHandCursor);
        } else {
            updateCursorForHandle(hitTestHandle(event->pos()));
        }
    }
}

void ChartWidget::mouseReleaseEvent(QMouseEvent*) {
    m_dragging = false;
    m_resizing = false;
    m_activeHandle = None;
    m_autoScrollTimer->stop();
    setCursor(Qt::ArrowCursor);
}

void ChartWidget::onAutoScroll() {
    if (!m_dragging || !parentWidget()) return;

    // parentWidget() is the viewport; its parent is the QAbstractScrollArea (QTableView)
    auto* scrollArea = qobject_cast<QAbstractScrollArea*>(parentWidget()->parentWidget());
    if (!scrollArea) return;

    // Determine scroll direction from where the unclamped position would be
    QPoint rawPos = m_lastGlobalPos - m_dragOffset;
    int maxX = parentWidget()->width() - width();
    int maxY = parentWidget()->height() - height();

    // QTableView uses scroll-per-item: scroll 5 rows / 2 cols per 16ms tick
    // (5 rows * 60 ticks/sec = ~300 rows/sec — very fast)
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

    // Chart stays at the viewport edge while the grid scrolls underneath.
    // Its effective grid position advances with the scroll.
    // Emit chartMoved so overlays (e.g. Data2App) reposition.
    emit chartMoved(this);
}

void ChartWidget::mouseDoubleClickEvent(QMouseEvent*) {
    emit propertiesRequested(this);
}

void ChartWidget::contextMenuEvent(QContextMenuEvent* event) {
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

    menu.addAction("Edit Chart...", this, [this]() { emit propertiesRequested(this); });
    menu.addAction("Refresh Data", this, [this]() { refreshData(); });
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
    menu.addAction("Delete Chart", this, [this]() { emit deleteRequested(this); });

    menu.exec(event->globalPos());
}

void ChartWidget::keyPressEvent(QKeyEvent* event) {
    if (m_selected && (event->key() == Qt::Key_Delete || event->key() == Qt::Key_Backspace)) {
        emit deleteRequested(this);
        event->accept();
        return;
    }
    QWidget::keyPressEvent(event);
}
