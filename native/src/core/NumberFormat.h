#ifndef NUMBERFORMAT_H
#define NUMBERFORMAT_H

#include <QString>
#include <QColor>

enum class NumberFormatType {
    General, Number, Currency, Accounting, Percentage,
    Date, Time, Fraction, Scientific, Text, Custom, Special
};

// Negative number display styles
enum class NegativeStyle {
    Minus,              // -1234.56
    Red,                // 1234.56 (in red)
    Parentheses,        // (1234.56)
    RedParentheses      // (1234.56) (in red)
};

struct NumberFormatOptions {
    NumberFormatType type = NumberFormatType::General;
    int decimalPlaces = 2;
    bool useThousandsSeparator = false;
    QString currencyCode = "USD";
    QString dateFormatId = "mm/dd/yyyy";
    QString customFormat = "";
    NegativeStyle negativeStyle = NegativeStyle::Minus;
    QString specialType = "";  // "zipcode", "zipcode4", "phone", "ssn"
};

// Result of formatting: includes text + optional color override
struct FormatResult {
    QString text;
    QColor color;       // Invalid = no color override (use default)
    bool hasColor = false;
};

struct CurrencyDef {
    QString code;
    QString symbol;
    QString label;
};

class NumberFormat {
public:
    // Format with full result (text + color)
    static FormatResult formatFull(const QString& value, const NumberFormatOptions& options);

    // Simple format (text only, backward compatible)
    static QString format(const QString& value, const NumberFormatOptions& options);

    // Custom format code parser (Excel-compatible)
    static FormatResult applyCustomFormatFull(const QString& value, const QString& formatStr);
    static QString applyCustomFormat(const QString& value, const QString& formatStr);

    static NumberFormatType typeFromString(const QString& str);
    static QString typeToString(NumberFormatType type);
    static QString getCurrencySymbol(const QString& code);

    static const QList<CurrencyDef>& currencies();
};

#endif // NUMBERFORMAT_H
