#include "ChartWidget.h"
#include "../core/Spreadsheet.h"
#include <QPainter>
#include <QPainterPath>
#include <QMouseEvent>
#include <QMenu>
#include <QtMath>
#include <algorithm>
#include <cmath>

// --- Theme color palettes ---
static const QVector<QVector<QColor>> kThemePalettes = {
    // 0: Excel
    { QColor("#4472C4"), QColor("#ED7D31"), QColor("#A5A5A5"), QColor("#FFC000"), QColor("#5B9BD5"), QColor("#70AD47") },
    // 1: Material
    { QColor("#2196F3"), QColor("#FF5722"), QColor("#4CAF50"), QColor("#FFC107"), QColor("#9C27B0"), QColor("#00BCD4") },
    // 2: Solarized
    { QColor("#268BD2"), QColor("#DC322F"), QColor("#859900"), QColor("#B58900"), QColor("#6C71C4"), QColor("#2AA198") },
    // 3: Dark
    { QColor("#00C8FF"), QColor("#FF6384"), QColor("#36A2EB"), QColor("#FFCE56"), QColor("#9966FF"), QColor("#FF9F40") },
    // 4: Monochrome
    { QColor("#333333"), QColor("#666666"), QColor("#999999"), QColor("#BBBBBB"), QColor("#444444"), QColor("#777777") },
    // 5: Pastel
    { QColor("#A8D8EA"), QColor("#FFB7B2"), QColor("#B5EAD7"), QColor("#FFDAC1"), QColor("#C7CEEA"), QColor("#E2F0CB") },
};

ChartWidget::ChartWidget(QWidget* parent)
    : QWidget(parent) {
    setMinimumSize(200, 150);
    resize(400, 300);
    setAttribute(Qt::WA_DeleteOnClose, false);
    setMouseTracking(true);

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

QVector<QColor> ChartWidget::getThemeColors() const {
    int idx = qBound(0, m_config.themeIndex, static_cast<int>(kThemePalettes.size()) - 1);
    return kThemePalettes[idx];
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

void ChartWidget::loadDataFromRange(const QString& range) {
    if (!m_spreadsheet || range.isEmpty()) return;

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

    for (int r = startRow + 1; r <= endRow; ++r) {
        auto val = m_spreadsheet->getCellValue(CellAddress(r, startCol));
        QString text = val.toString();
        bool ok;
        double num = text.toDouble(&ok);
        xValues.append(ok ? num : static_cast<double>(r - startRow));
    }

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

    bool first = true;
    for (const auto& s : m_config.series) {
        for (double v : s.yValues) {
            if (first) {
                minVal = maxVal = v;
                first = false;
            } else {
                minVal = qMin(minVal, v);
                maxVal = qMax(maxVal, v);
            }
        }
    }

    if (minVal == maxVal) {
        minVal -= 1;
        maxVal += 1;
    }

    // Line charts: dynamic y-axis from data range (better for showing trends)
    // All other charts: always start at 0 (accurate visual comparison of values)
    bool isLineChart = (m_config.type == ChartType::Line);
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
    int right = width() - 15;
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
    }

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

    // White background with rounded corners
    p.setBrush(Qt::white);
    p.setPen(QPen(QColor("#D0D5DD"), 1));
    p.drawRoundedRect(area.adjusted(0, 0, -1, -1), 12, 12);

    // Clip to rounded rect so content doesn't overflow corners
    QPainterPath clipPath;
    clipPath.addRoundedRect(area.adjusted(1, 1, -2, -2), 11, 11);
    p.setClipPath(clipPath);
}

void ChartWidget::drawTitle(QPainter& p, const QRect& area) {
    if (m_config.title.isEmpty()) return;
    p.setPen(QColor("#333"));
    p.setFont(QFont("Arial", 13, QFont::Bold));
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
        if (std::abs(v) >= 1000000) label = QString::number(v / 1000000.0, 'f', 1) + "M";
        else if (std::abs(v) >= 1000) label = QString::number(v / 1000.0, 'f', 1) + "K";
        else label = QString::number(v, 'f', step < 1 ? 1 : 0);

        p.drawText(QRect(plotArea.left() - AXIS_MARGIN, y - 8, AXIS_MARGIN - 6, 16),
                   Qt::AlignRight | Qt::AlignVCenter, label);
    }

    // X axis category labels
    if (!m_config.series.isEmpty() && !m_config.series[0].xValues.isEmpty()) {
        int n = m_config.series[0].xValues.size();
        int maxLabels = qMax(1, plotArea.width() / 50);
        int labelStep = qMax(1, n / maxLabels);

        for (int i = 0; i < n; i += labelStep) {
            double frac = (n > 1) ? static_cast<double>(i) / (n - 1) : 0.5;
            int x = plotArea.left() + static_cast<int>(frac * plotArea.width());
            p.drawLine(x, plotArea.bottom(), x, plotArea.bottom() + 4);
            p.drawText(QRect(x - 25, plotArea.bottom() + 5, 50, 16),
                       Qt::AlignCenter, QString::number(i + 1));
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

void ChartWidget::drawLegend(QPainter& p, const QRect& area) {
    if (m_config.series.isEmpty()) return;

    p.setFont(QFont("Arial", 9));
    int y = area.bottom() - LEGEND_HEIGHT;
    int totalWidth = 0;

    for (const auto& s : m_config.series) {
        totalWidth += 14 + p.fontMetrics().horizontalAdvance(s.name) + 16;
    }

    int x = (area.width() - totalWidth) / 2;

    for (const auto& s : m_config.series) {
        p.setPen(Qt::NoPen);
        p.setBrush(s.color);
        p.drawRoundedRect(x, y + 4, 10, 10, 2, 2);
        x += 14;

        p.setPen(QColor("#555"));
        p.drawText(x, y + 13, s.name);
        x += p.fontMetrics().horizontalAdvance(s.name) + 16;
    }
}

void ChartWidget::drawSelectionHandles(QPainter& p) {
    p.setPen(QPen(QColor("#4A90D9"), 2));
    p.setBrush(Qt::NoBrush);
    p.drawRect(rect().adjusted(1, 1, -2, -2));

    // Corner and edge handles
    p.setPen(QPen(QColor("#4A90D9"), 1));
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

    for (const auto& s : m_config.series) {
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
    double barWidth = (groupWidth * 0.7) / numSeries;
    double gap = groupWidth * 0.15;

    for (int si = 0; si < numSeries; ++si) {
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

// --- Bar Chart (horizontal) ---

void ChartWidget::drawBarChart(QPainter& p, const QRect& plotArea) {
    if (m_config.series.isEmpty()) return;

    double minVal, maxVal, step;
    computeAxisRange(minVal, maxVal, step);

    int numSeries = m_config.series.size();
    int numPoints = m_config.series[0].yValues.size();
    if (numPoints == 0) return;

    double groupHeight = static_cast<double>(plotArea.height()) / numPoints;
    double barHeight = (groupHeight * 0.7) / numSeries;
    double gap = groupHeight * 0.15;

    for (int si = 0; si < numSeries; ++si) {
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

// --- Scatter Chart ---

void ChartWidget::drawScatterChart(QPainter& p, const QRect& plotArea) {
    if (m_config.series.isEmpty()) return;

    double minVal, maxVal, step;
    computeAxisRange(minVal, maxVal, step);

    // Compute X range
    double xMin = 0, xMax = 1;
    bool first = true;
    for (const auto& s : m_config.series) {
        for (double v : s.xValues) {
            if (first) { xMin = xMax = v; first = false; }
            else { xMin = qMin(xMin, v); xMax = qMax(xMax, v); }
        }
    }
    if (xMin == xMax) { xMin -= 1; xMax += 1; }

    int pointRadius = qMax(1, static_cast<int>(5 * m_animProgress));

    for (const auto& s : m_config.series) {
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
    for (double v : s.yValues) total += qMax(0.0, v);
    if (total <= 0) return;

    auto colors = getThemeColors();
    int size = qMin(plotArea.width(), plotArea.height()) - 20;
    QRect pieRect(plotArea.center().x() - size / 2, plotArea.center().y() - size / 2, size, size);

    int startAngle = 90 * 16; // Start from top
    for (int i = 0; i < s.yValues.size(); ++i) {
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

    for (int si = m_config.series.size() - 1; si >= 0; --si) {
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

// --- Donut Chart ---

void ChartWidget::drawDonutChart(QPainter& p, const QRect& plotArea) {
    if (m_config.series.isEmpty() || m_config.series[0].yValues.isEmpty()) return;

    const auto& s = m_config.series[0];
    double total = 0;
    for (double v : s.yValues) total += qMax(0.0, v);
    if (total <= 0) return;

    auto colors = getThemeColors();
    int size = qMin(plotArea.width(), plotArea.height()) - 20;
    QRect outerRect(plotArea.center().x() - size / 2, plotArea.center().y() - size / 2, size, size);
    int innerSize = size * 55 / 100;
    QRect innerRect(plotArea.center().x() - innerSize / 2, plotArea.center().y() - innerSize / 2,
                    innerSize, innerSize);

    int startAngle = 90 * 16;
    for (int i = 0; i < s.yValues.size(); ++i) {
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
        QPoint newPos = event->globalPosition().toPoint() - m_dragOffset;
        // Constrain to parent
        if (parentWidget()) {
            newPos.setX(qBound(0, newPos.x(), parentWidget()->width() - width()));
            newPos.setY(qBound(0, newPos.y(), parentWidget()->height() - height()));
        }
        move(newPos);
        emit chartMoved(this);
    } else {
        updateCursorForHandle(hitTestHandle(event->pos()));
    }
}

void ChartWidget::mouseReleaseEvent(QMouseEvent*) {
    m_dragging = false;
    m_resizing = false;
    m_activeHandle = None;
    setCursor(Qt::ArrowCursor);
}

void ChartWidget::mouseDoubleClickEvent(QMouseEvent*) {
    emit propertiesRequested(this);
}

void ChartWidget::contextMenuEvent(QContextMenuEvent* event) {
    QMenu menu(this);
    menu.setStyleSheet(
        "QMenu { background: #FFFFFF; border: 1px solid #D0D5DD; border-radius: 6px; padding: 4px; }"
        "QMenu::item { padding: 6px 20px; border-radius: 4px; }"
        "QMenu::item:selected { background-color: #E8F0FE; }"
        "QMenu::separator { height: 1px; background: #E0E3E8; margin: 3px 8px; }"
    );

    menu.addAction("Edit Chart...", this, [this]() { emit propertiesRequested(this); });
    menu.addAction("Refresh Data", this, [this]() { refreshData(); });
    menu.addSeparator();
    menu.addAction("Delete Chart", this, [this]() { emit deleteRequested(this); });

    menu.exec(event->globalPos());
}
