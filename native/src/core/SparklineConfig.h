#ifndef SPARKLINECONFIG_H
#define SPARKLINECONFIG_H

#include <QString>
#include <QColor>
#include <QVector>
#include <QMetaType>

enum class SparklineType {
    Line,
    Column,
    WinLoss
};

struct SparklineConfig {
    SparklineType type = SparklineType::Line;
    QString dataRange;              // e.g., "B2:G2"
    QColor lineColor = QColor("#4472C4");
    QColor highPointColor = QColor("#22C55E");
    QColor lowPointColor = QColor("#EF4444");
    QColor negativeColor = QColor("#EF4444");
    bool showHighPoint = false;
    bool showLowPoint = false;
    int lineWidth = 2;
};

struct SparklineRenderData {
    SparklineType type = SparklineType::Line;
    QVector<double> values;
    double minVal = 0;
    double maxVal = 0;
    QColor lineColor;
    QColor highPointColor;
    QColor lowPointColor;
    QColor negativeColor;
    bool showHighPoint = false;
    bool showLowPoint = false;
    int lineWidth = 2;
    int highIndex = -1;
    int lowIndex = -1;
};

Q_DECLARE_METATYPE(SparklineRenderData)

#endif // SPARKLINECONFIG_H
