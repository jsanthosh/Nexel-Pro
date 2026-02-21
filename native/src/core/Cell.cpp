#include "Cell.h"
#include <QDateTime>

Cell::Cell() : m_type(CellType::Empty), m_dirty(false) {
    // m_customStyle is nullptr by default — uses shared static defaultStyle()
}

const CellStyle& Cell::defaultStyle() {
    static const CellStyle s_default;
    return s_default;
}

void Cell::setValue(const QVariant& value) {
    if (m_value != value) {
        m_value = value;
        m_dirty = true;

        // Detect type
        if (value.isNull() || !value.isValid()) {
            m_type = CellType::Empty;
        } else if (value.type() == QVariant::Bool) {
            m_type = CellType::Boolean;
        } else if (value.type() == QVariant::Int || value.type() == QVariant::Double) {
            m_type = CellType::Number;
        } else if (value.type() == QVariant::Date || value.type() == QVariant::DateTime) {
            m_type = CellType::Date;
        } else {
            m_type = CellType::Text;
        }
    }
}

void Cell::setFormula(const QString& formula) {
    if (m_formula != formula) {
        m_formula = formula;
        m_type = CellType::Formula;
        m_dirty = true;
    }
}

QVariant Cell::getValue() const {
    return m_value;
}

QString Cell::getFormula() const {
    return m_formula;
}

CellType Cell::getType() const {
    return m_type;
}

void Cell::setStyle(const CellStyle& style) {
    m_customStyle = std::make_unique<CellStyle>(style);
}

const CellStyle& Cell::getStyle() const {
    return m_customStyle ? *m_customStyle : defaultStyle();
}

void Cell::setComputedValue(const QVariant& value) {
    m_computedValue = value;
}

QVariant Cell::getComputedValue() const {
    return m_computedValue;
}

bool Cell::isDirty() const {
    return m_dirty;
}

void Cell::setDirty(bool dirty) {
    m_dirty = dirty;
}

bool Cell::hasError() const {
    return m_type == CellType::Error;
}

void Cell::setError(const QString& error) {
    m_error = error;
    m_type = CellType::Error;
}

QString Cell::getError() const {
    return m_error;
}

QString Cell::toString() const {
    switch (m_type) {
        case CellType::Formula:
            return m_formula;
        case CellType::Date:
            return m_value.toDateTime().toString(Qt::ISODate);
        case CellType::Number:
            return m_value.toString();
        case CellType::Boolean:
            return m_value.toBool() ? "TRUE" : "FALSE";
        case CellType::Error:
            return "#" + m_error;
        case CellType::Text:
        case CellType::Empty:
        default:
            return m_value.toString();
    }
}

void Cell::clear() {
    m_value = QVariant();
    m_formula = QString();
    m_computedValue = QVariant();
    m_type = CellType::Empty;
    m_customStyle.reset();
    m_error = QString();
    m_dirty = true;
}
