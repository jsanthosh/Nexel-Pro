#ifndef CONDITIONALFORMATTING_H
#define CONDITIONALFORMATTING_H

#include <QString>
#include <QVariant>
#include <QColor>
#include <vector>
#include <memory>
#include <functional>
#include <optional>
#include "CellRange.h"
#include "Cell.h"

enum class ConditionType {
    Equal,
    NotEqual,
    GreaterThan,
    LessThan,
    GreaterThanOrEqual,
    LessThanOrEqual,
    Between,
    CellContains,
    Formula,
    // Visual formatting types
    DataBar,
    ColorScale2,
    ColorScale3,
    IconSet
};

// --- Visual formatting configuration structs ---

struct DataBarConfig {
    QColor barColor = QColor(99, 142, 198); // default blue
    double minValue = 0;
    double maxValue = 100;
    bool autoRange = true;  // auto-detect min/max from data
    bool showValue = true;  // show number alongside bar
};

struct ColorScaleConfig {
    QColor minColor = QColor(248, 105, 107); // red
    QColor midColor = QColor(255, 235, 132); // yellow (3-color only)
    QColor maxColor = QColor(99, 190, 123);  // green
    double minValue = 0;
    double maxValue = 100;
    bool autoRange = true;
    bool threeColor = false;
};

struct IconSetConfig {
    enum IconType { TrafficLights, Arrows3, Flags3, Stars3, Checkmarks3 };
    IconType iconType = TrafficLights;
    // Thresholds as percentages (0-100)
    double threshold1 = 33.33; // below this = icon1 (red/down)
    double threshold2 = 66.67; // below this = icon2 (yellow/neutral), above = icon3 (green/up)
    bool showValue = true;
    bool reverseOrder = false;
};

// Result of evaluating visual formatting for a cell
struct VisualFormatResult {
    ConditionType type = ConditionType::Equal; // which visual type
    // DataBar
    double barFraction = 0.0; // 0..1, width fraction
    QColor barColor;
    bool barShowValue = true;
    // ColorScale
    QColor scaleColor;
    // IconSet
    int iconIndex = -1; // 0=low, 1=mid, 2=high
    IconSetConfig::IconType iconType = IconSetConfig::TrafficLights;
    bool iconShowValue = true;
    bool reverseOrder = false;
};

class ConditionalFormat {
public:
    ConditionalFormat(const CellRange& range, ConditionType type);

    const CellRange& getRange() const;
    ConditionType getType() const;
    const CellStyle& getStyle() const;
    const QVariant& getValue1() const { return m_value1; }
    const QVariant& getValue2() const { return m_value2; }
    const QString& getFormula() const { return m_formula; }

    void setValue1(const QVariant& value);
    void setValue2(const QVariant& value);
    void setFormula(const QString& formula);
    void setStyle(const CellStyle& style);

    // Visual formatting configs
    void setDataBarConfig(const DataBarConfig& cfg) { m_dataBarConfig = cfg; }
    const DataBarConfig& getDataBarConfig() const { return m_dataBarConfig; }

    void setColorScaleConfig(const ColorScaleConfig& cfg) { m_colorScaleConfig = cfg; }
    const ColorScaleConfig& getColorScaleConfig() const { return m_colorScaleConfig; }

    void setIconSetConfig(const IconSetConfig& cfg) { m_iconSetConfig = cfg; }
    const IconSetConfig& getIconSetConfig() const { return m_iconSetConfig; }

    bool isVisualType() const;

    using FormulaEvaluator = std::function<QVariant(const QString&)>;
    bool matches(const QVariant& cellValue, const FormulaEvaluator& evaluator = nullptr) const;

private:
    CellRange m_range;
    ConditionType m_type;
    QVariant m_value1;
    QVariant m_value2;
    QString m_formula;
    CellStyle m_style;
    DataBarConfig m_dataBarConfig;
    ColorScaleConfig m_colorScaleConfig;
    IconSetConfig m_iconSetConfig;
};

class ConditionalFormatting {
public:
    ConditionalFormatting() = default;
    ~ConditionalFormatting() = default;

    // Add formatting rule
    void addRule(std::shared_ptr<ConditionalFormat> rule);

    // Remove rule
    void removeRule(size_t index);

    // Get rules for a specific range
    std::vector<std::shared_ptr<ConditionalFormat>> getRulesForRange(const CellRange& range) const;

    // Set formula evaluator for formula-based conditions
    using FormulaEvaluator = ConditionalFormat::FormulaEvaluator;
    void setFormulaEvaluator(const FormulaEvaluator& evaluator) { m_evaluator = evaluator; }

    // Get style for a cell
    CellStyle getEffectiveStyle(const CellAddress& addr, const QVariant& cellValue, const CellStyle& baseStyle) const;

    // Get visual formatting result for a cell (data bar, color scale, icon set)
    using ValueLookup = std::function<QVariant(int row, int col)>;
    std::optional<VisualFormatResult> getVisualFormat(const CellAddress& addr, const QVariant& cellValue,
                                                       const ValueLookup& valueLookup = nullptr) const;

    // Get all rules
    const std::vector<std::shared_ptr<ConditionalFormat>>& getAllRules() const;

    // Clear all rules
    void clearRules();

private:
    std::vector<std::shared_ptr<ConditionalFormat>> m_rules;
    FormulaEvaluator m_evaluator;

    // Helper: compute min/max across a range for auto-range
    static std::pair<double, double> computeRangeMinMax(const CellRange& range, const ValueLookup& valueLookup);
};

#endif // CONDITIONALFORMATTING_H
