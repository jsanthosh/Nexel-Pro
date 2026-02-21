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
    CountDistinct
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

    void addValue(double val, const QString& rawVal);
    double result(AggregationFunction func) const;
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
