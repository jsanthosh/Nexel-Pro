#include "ConditionalFormatting.h"
#include <algorithm>
#include <cmath>

ConditionalFormat::ConditionalFormat(const CellRange& range, ConditionType type)
    : m_range(range), m_type(type) {
}

const CellRange& ConditionalFormat::getRange() const {
    return m_range;
}

ConditionType ConditionalFormat::getType() const {
    return m_type;
}

const CellStyle& ConditionalFormat::getStyle() const {
    return m_style;
}

void ConditionalFormat::setValue1(const QVariant& value) {
    m_value1 = value;
}

void ConditionalFormat::setValue2(const QVariant& value) {
    m_value2 = value;
}

void ConditionalFormat::setFormula(const QString& formula) {
    m_formula = formula;
}

void ConditionalFormat::setStyle(const CellStyle& style) {
    m_style = style;
}

bool ConditionalFormat::isVisualType() const {
    return m_type == ConditionType::DataBar ||
           m_type == ConditionType::ColorScale2 ||
           m_type == ConditionType::ColorScale3 ||
           m_type == ConditionType::IconSet;
}

bool ConditionalFormat::matches(const QVariant& cellValue, const FormulaEvaluator& evaluator) const {
    // Visual formatting types always "match" — they apply to all numeric cells in the range
    if (isVisualType()) return true;

    switch (m_type) {
        case ConditionType::Equal:
            return cellValue == m_value1;
        case ConditionType::NotEqual:
            return cellValue != m_value1;
        case ConditionType::GreaterThan:
            return cellValue.toDouble() > m_value1.toDouble();
        case ConditionType::LessThan:
            return cellValue.toDouble() < m_value1.toDouble();
        case ConditionType::GreaterThanOrEqual:
            return cellValue.toDouble() >= m_value1.toDouble();
        case ConditionType::LessThanOrEqual:
            return cellValue.toDouble() <= m_value1.toDouble();
        case ConditionType::Between:
            return cellValue.toDouble() >= m_value1.toDouble() &&
                   cellValue.toDouble() <= m_value2.toDouble();
        case ConditionType::CellContains:
            return cellValue.toString().contains(m_value1.toString());
        case ConditionType::Formula:
            if (evaluator && !m_formula.isEmpty()) {
                QVariant result = evaluator(m_formula);
                if (result.typeId() == QMetaType::Bool) return result.toBool();
                if (result.typeId() == QMetaType::Double || result.typeId() == QMetaType::Int)
                    return result.toDouble() != 0.0;
                return !result.toString().isEmpty() && result.toString().toLower() != "false";
            }
            return false;
        case ConditionType::TopN:
        case ConditionType::TopNPercent:
        case ConditionType::BottomN:
        case ConditionType::BottomNPercent:
        case ConditionType::AboveAverage:
        case ConditionType::BelowAverage:
            // These require range-level evaluation (handled in getEffectiveStyle)
            return false;
        case ConditionType::DuplicateValues:
        case ConditionType::UniqueValues:
            // Requires scanning the range for duplicates
            return false;
        case ConditionType::DateOccurring:
            return false;
        default:
            return false;
    }
    return false;
}

void ConditionalFormatting::addRule(std::shared_ptr<ConditionalFormat> rule) {
    m_rules.push_back(rule);
}

void ConditionalFormatting::removeRule(size_t index) {
    if (index < m_rules.size()) {
        m_rules.erase(m_rules.begin() + index);
    }
}

std::vector<std::shared_ptr<ConditionalFormat>> ConditionalFormatting::getRulesForRange(const CellRange& range) const {
    std::vector<std::shared_ptr<ConditionalFormat>> result;
    for (const auto& rule : m_rules) {
        if (rule->getRange().intersects(range)) {
            result.push_back(rule);
        }
    }
    return result;
}

CellStyle ConditionalFormatting::getEffectiveStyle(const CellAddress& addr, const QVariant& cellValue, const CellStyle& baseStyle) const {
    CellStyle effective = baseStyle;

    for (const auto& rule : m_rules) {
        // Skip visual formatting types — they don't modify CellStyle
        if (rule->isVisualType()) continue;

        if (rule->getRange().contains(addr) && rule->matches(cellValue, m_evaluator)) {
            const CellStyle& ruleStyle = rule->getStyle();
            if (ruleStyle.bold) effective.bold = true;
            if (ruleStyle.italic) effective.italic = true;
            if (ruleStyle.underline) effective.underline = true;
            if (ruleStyle.foregroundColor != "#000000") effective.foregroundColor = ruleStyle.foregroundColor;
            if (ruleStyle.backgroundColor != "#FFFFFF") effective.backgroundColor = ruleStyle.backgroundColor;
            if (ruleStyle.fontName != "Arial") effective.fontName = ruleStyle.fontName;
            if (ruleStyle.fontSize != 11) effective.fontSize = ruleStyle.fontSize;
        }
    }

    return effective;
}

// Helper to interpolate between two colors
static QColor interpolateColor(const QColor& c1, const QColor& c2, double t) {
    t = std::clamp(t, 0.0, 1.0);
    return QColor(
        static_cast<int>(c1.red()   + (c2.red()   - c1.red())   * t),
        static_cast<int>(c1.green() + (c2.green() - c1.green()) * t),
        static_cast<int>(c1.blue()  + (c2.blue()  - c1.blue())  * t)
    );
}

std::pair<double, double> ConditionalFormatting::computeRangeMinMax(const CellRange& range, const ValueLookup& valueLookup) {
    double minVal = std::numeric_limits<double>::max();
    double maxVal = std::numeric_limits<double>::lowest();
    bool found = false;

    if (!valueLookup) return {0.0, 100.0};

    auto start = range.getStart();
    auto end = range.getEnd();
    for (int r = start.row; r <= end.row; ++r) {
        for (int c = start.col; c <= end.col; ++c) {
            QVariant v = valueLookup(r, c);
            bool ok = false;
            double d = v.toDouble(&ok);
            if (ok) {
                minVal = std::min(minVal, d);
                maxVal = std::max(maxVal, d);
                found = true;
            }
        }
    }
    if (!found) return {0.0, 100.0};
    if (minVal == maxVal) { minVal -= 1.0; maxVal += 1.0; }
    return {minVal, maxVal};
}

std::optional<VisualFormatResult> ConditionalFormatting::getVisualFormat(
    const CellAddress& addr, const QVariant& cellValue, const ValueLookup& valueLookup) const {

    bool ok = false;
    double numValue = cellValue.toDouble(&ok);
    if (!ok) return std::nullopt; // visual formatting only applies to numeric cells

    for (const auto& rule : m_rules) {
        if (!rule->isVisualType()) continue;
        if (!rule->getRange().contains(addr)) continue;

        VisualFormatResult result;
        result.type = rule->getType();

        double minVal, maxVal;

        switch (rule->getType()) {
        case ConditionType::DataBar: {
            const auto& cfg = rule->getDataBarConfig();
            if (cfg.autoRange) {
                auto [lo, hi] = computeRangeMinMax(rule->getRange(), valueLookup);
                minVal = lo;
                maxVal = hi;
            } else {
                minVal = cfg.minValue;
                maxVal = cfg.maxValue;
            }
            double range = maxVal - minVal;
            if (range <= 0) range = 1.0;
            result.barFraction = std::clamp((numValue - minVal) / range, 0.0, 1.0);
            result.barColor = cfg.barColor;
            result.barShowValue = cfg.showValue;
            return result;
        }
        case ConditionType::ColorScale2: {
            const auto& cfg = rule->getColorScaleConfig();
            if (cfg.autoRange) {
                auto [lo, hi] = computeRangeMinMax(rule->getRange(), valueLookup);
                minVal = lo;
                maxVal = hi;
            } else {
                minVal = cfg.minValue;
                maxVal = cfg.maxValue;
            }
            double range = maxVal - minVal;
            if (range <= 0) range = 1.0;
            double t = std::clamp((numValue - minVal) / range, 0.0, 1.0);
            result.scaleColor = interpolateColor(cfg.minColor, cfg.maxColor, t);
            return result;
        }
        case ConditionType::ColorScale3: {
            const auto& cfg = rule->getColorScaleConfig();
            if (cfg.autoRange) {
                auto [lo, hi] = computeRangeMinMax(rule->getRange(), valueLookup);
                minVal = lo;
                maxVal = hi;
            } else {
                minVal = cfg.minValue;
                maxVal = cfg.maxValue;
            }
            double range = maxVal - minVal;
            if (range <= 0) range = 1.0;
            double t = std::clamp((numValue - minVal) / range, 0.0, 1.0);
            if (t <= 0.5) {
                // Interpolate minColor -> midColor
                result.scaleColor = interpolateColor(cfg.minColor, cfg.midColor, t * 2.0);
            } else {
                // Interpolate midColor -> maxColor
                result.scaleColor = interpolateColor(cfg.midColor, cfg.maxColor, (t - 0.5) * 2.0);
            }
            return result;
        }
        case ConditionType::IconSet: {
            const auto& cfg = rule->getIconSetConfig();
            {
                auto [lo, hi] = computeRangeMinMax(rule->getRange(), valueLookup);
                minVal = lo;
                maxVal = hi;
            }
            double range = maxVal - minVal;
            if (range <= 0) range = 1.0;
            double pct = std::clamp((numValue - minVal) / range * 100.0, 0.0, 100.0);
            if (pct < cfg.threshold1) {
                result.iconIndex = 0; // low (red/down)
            } else if (pct < cfg.threshold2) {
                result.iconIndex = 1; // mid (yellow/neutral)
            } else {
                result.iconIndex = 2; // high (green/up)
            }
            if (cfg.reverseOrder) {
                result.iconIndex = 2 - result.iconIndex;
            }
            result.iconType = cfg.iconType;
            result.iconShowValue = cfg.showValue;
            result.reverseOrder = cfg.reverseOrder;
            return result;
        }
        default:
            break;
        }
    }

    return std::nullopt;
}

const std::vector<std::shared_ptr<ConditionalFormat>>& ConditionalFormatting::getAllRules() const {
    return m_rules;
}

void ConditionalFormatting::clearRules() {
    m_rules.clear();
}
