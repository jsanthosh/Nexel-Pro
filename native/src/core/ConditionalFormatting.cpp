#include "ConditionalFormatting.h"

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

bool ConditionalFormat::matches(const QVariant& cellValue, const FormulaEvaluator& evaluator) const {
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

const std::vector<std::shared_ptr<ConditionalFormat>>& ConditionalFormatting::getAllRules() const {
    return m_rules;
}

void ConditionalFormatting::clearRules() {
    m_rules.clear();
}
