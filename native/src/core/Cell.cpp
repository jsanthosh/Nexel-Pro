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
    // Detect type
    if (value.isNull() || !value.isValid()) {
        if (m_type != CellType::Empty) {
            m_value = value;
            m_type = CellType::Empty;
            m_dirty = true;
        }
    } else if (value.type() == QVariant::Bool) {
        m_value = value;
        m_type = CellType::Boolean;
        m_dirty = true;
    } else if (value.type() == QVariant::Int || value.type() == QVariant::Double) {
        m_value = value;
        m_type = CellType::Number;
        m_dirty = true;
    } else if (value.type() == QVariant::Date || value.type() == QVariant::DateTime) {
        m_value = value;
        m_type = CellType::Date;
        m_dirty = true;
    } else {
        // String value — auto-detect numbers (like Excel)
        QString str = value.toString();
        if (str.isEmpty()) {
            m_value = value;
            m_type = CellType::Empty;
            m_dirty = true;
        } else {
            bool ok;
            double num = str.toDouble(&ok);
            if (ok) {
                m_value = QVariant(num);
                m_type = CellType::Number;
            } else {
                m_value = value;
                m_type = CellType::Text;
            }
            m_dirty = true;
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
    // Skip allocation if style matches default (common for bulk import with styleIdx 0)
    const auto& def = defaultStyle();
    if (style.fontName == def.fontName && style.fontSize == def.fontSize &&
        style.bold == def.bold && style.italic == def.italic &&
        style.underline == def.underline && style.strikethrough == def.strikethrough &&
        style.foregroundColor == def.foregroundColor &&
        style.backgroundColor == def.backgroundColor &&
        style.numberFormat == def.numberFormat &&
        !style.borderTop.enabled && !style.borderBottom.enabled &&
        !style.borderLeft.enabled && !style.borderRight.enabled &&
        style.hAlign == def.hAlign && style.vAlign == def.vAlign &&
        style.textOverflow == def.textOverflow) {
        m_customStyle.reset(); // use default
        return;
    }
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

QString Cell::getComment() const {
    return m_comment;
}

void Cell::setComment(const QString& comment) {
    m_comment = comment;
}

bool Cell::hasComment() const {
    return !m_comment.isEmpty();
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
    m_hyperlink = QString();
    m_dirty = true;
    m_spillParent = CellAddress(-1, -1);
}
