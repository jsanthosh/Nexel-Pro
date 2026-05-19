#ifndef PIVOTENGINE_H
#define PIVOTENGINE_H

#include <QString>
#include <QVariant>
#include <QStringList>
#include <vector>
#include <map>
#include <set>
#include <memory>
#include <limits>
#include "CellRange.h"

class Spreadsheet;

enum class AggregationFunction {
    Sum,
    Count,
    Average,
    Min,
    Max,
    CountDistinct,
    // Statistical aggregations (round 2)
    StDev,    // Sample standard deviation (n-1 denominator)
    StDevP,   // Population standard deviation
    Var,      // Sample variance
    VarP,     // Population variance
    Median,
    Product
};

struct PivotField {
    int sourceColumnIndex = 0;
    QString name;
};

struct PivotValueField {
    int sourceColumnIndex = 0;
    QString name;
    AggregationFunction aggregation = AggregationFunction::Sum;

    QString displayName() const;
};

struct PivotFilterField {
    int sourceColumnIndex = 0;
    QString name;
    QStringList selectedValues; // empty = all
};

struct PivotConfig {
    CellRange sourceRange;
    int sourceSheetIndex = 0;

    std::vector<PivotField> rowFields;
    std::vector<PivotField> columnFields;
    std::vector<PivotValueField> valueFields;
    std::vector<PivotFilterField> filterFields;

    bool showGrandTotalRow = true;
    bool showGrandTotalColumn = true;
    bool showSubtotals = true;
    bool autoChart = false;
    int chartType = 0; // 0=Column, maps to ChartType enum
};

struct AggregateAccumulator {
    double sum = 0.0;
    double min = std::numeric_limits<double>::max();
    double max = std::numeric_limits<double>::lowest();
    int count = 0;
    std::set<QString> distinctValues;
    // Online statistics: Welford's algorithm for mean + M2 (sum of squared
    // deviations from mean). Keeps numeric stability for big-N data without
    // storing every value. Median still needs the full sample, so we store
    // it separately and only when actually requested.
    double mean = 0.0;
    double m2 = 0.0;       // sum of (x - mean)^2 accumulated incrementally
    double product = 1.0;
    std::vector<double> samples; // populated only when Median is requested

    void addValue(double val, const QString& rawVal, bool needSamples = false);
    double result(AggregationFunction func);
};

struct PivotResult {
    std::vector<std::vector<QString>> rowLabels;
    std::vector<std::vector<QString>> columnLabels;
    std::vector<std::vector<QVariant>> data;

    std::vector<QVariant> grandTotalRow;
    std::vector<QVariant> grandTotalColumn;
    QVariant grandTotal;

    int numRowHeaderColumns = 0;
    int numColHeaderRows = 0;
    int dataStartRow = 0;
    int dataStartCol = 0;
};

class PivotEngine {
public:
    PivotEngine();

    void setSource(std::shared_ptr<Spreadsheet> sheet, const PivotConfig& config);
    PivotResult compute();
    void writeToSheet(std::shared_ptr<Spreadsheet> targetSheet,
                      const PivotResult& result,
                      const PivotConfig& config);

    QStringList getUniqueValues(std::shared_ptr<Spreadsheet> sheet,
                                const CellRange& range, int columnIndex);
    QStringList detectColumnHeaders(std::shared_ptr<Spreadsheet> sheet,
                                    const CellRange& range);

private:
    struct DataRow {
        std::vector<QVariant> values;
    };

    std::vector<DataRow> extractSourceData(std::shared_ptr<Spreadsheet> sheet,
                                           const CellRange& range,
                                           const std::vector<PivotFilterField>& filters);
    QString buildKey(const DataRow& row, const std::vector<PivotField>& fields);
    std::vector<QString> splitKey(const QString& key, int fieldCount);

    std::shared_ptr<Spreadsheet> m_sourceSheet;
    PivotConfig m_config;
};

#endif // PIVOTENGINE_H
