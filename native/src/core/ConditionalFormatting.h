#ifndef CONDITIONALFORMATTING_H
#define CONDITIONALFORMATTING_H

#include <QString>
#include <QVariant>
#include <vector>
#include <memory>
#include <functional>
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
    Formula
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

    using FormulaEvaluator = std::function<QVariant(const QString&)>;
    bool matches(const QVariant& cellValue, const FormulaEvaluator& evaluator = nullptr) const;

private:
    CellRange m_range;
    ConditionType m_type;
    QVariant m_value1;
    QVariant m_value2;
    QString m_formula;
    CellStyle m_style;
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

    // Get all rules
    const std::vector<std::shared_ptr<ConditionalFormat>>& getAllRules() const;

    // Clear all rules
    void clearRules();

private:
    std::vector<std::shared_ptr<ConditionalFormat>> m_rules;
    FormulaEvaluator m_evaluator;
};

#endif // CONDITIONALFORMATTING_H
