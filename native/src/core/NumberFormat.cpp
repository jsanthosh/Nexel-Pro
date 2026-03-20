#include "NumberFormat.h"
#include <QLocale>
#include <QDate>
#include <QTime>
#include <QDateTime>
#include <QRegularExpression>
#include <cmath>

// ============================================================================
// Currency definitions (50+ currencies)
// ============================================================================
static const QList<CurrencyDef> s_currencies = {
    {"USD", "$",       "US Dollar ($)"},
    {"EUR", "\u20AC",  "Euro (\u20AC)"},
    {"GBP", "\u00A3",  "British Pound (\u00A3)"},
    {"JPY", "\u00A5",  "Japanese Yen (\u00A5)"},
    {"INR", "\u20B9",  "Indian Rupee (\u20B9)"},
    {"CNY", "\u00A5",  "Chinese Yuan (\u00A5)"},
    {"KRW", "\u20A9",  "Korean Won (\u20A9)"},
    {"CAD", "CA$",     "Canadian Dollar (CA$)"},
    {"AUD", "A$",      "Australian Dollar (A$)"},
    {"CHF", "CHF",     "Swiss Franc (CHF)"},
    {"BRL", "R$",      "Brazilian Real (R$)"},
    {"MXN", "MX$",     "Mexican Peso (MX$)"},
    {"SGD", "S$",      "Singapore Dollar (S$)"},
    {"HKD", "HK$",     "Hong Kong Dollar (HK$)"},
    {"NOK", "kr",      "Norwegian Krone (kr)"},
    {"SEK", "kr",      "Swedish Krona (kr)"},
    {"DKK", "kr",      "Danish Krone (kr)"},
    {"NZD", "NZ$",     "New Zealand Dollar (NZ$)"},
    {"ZAR", "R",       "South African Rand (R)"},
    {"RUB", "\u20BD",  "Russian Ruble (\u20BD)"},
    {"TRY", "\u20BA",  "Turkish Lira (\u20BA)"},
    {"PLN", "z\u0142", "Polish Zloty (z\u0142)"},
    {"THB", "\u0E3F",  "Thai Baht (\u0E3F)"},
    {"IDR", "Rp",      "Indonesian Rupiah (Rp)"},
    {"MYR", "RM",      "Malaysian Ringgit (RM)"},
    {"PHP", "\u20B1",  "Philippine Peso (\u20B1)"},
    {"CZK", "K\u010D", "Czech Koruna (K\u010D)"},
    {"ILS", "\u20AA",  "Israeli Shekel (\u20AA)"},
    {"CLP", "CLP$",    "Chilean Peso (CLP$)"},
    {"AED", "AED",     "UAE Dirham (AED)"},
    {"SAR", "SAR",     "Saudi Riyal (SAR)"},
    {"TWD", "NT$",     "Taiwan Dollar (NT$)"},
    {"ARS", "AR$",     "Argentine Peso (AR$)"},
    {"COP", "COL$",    "Colombian Peso (COL$)"},
    {"EGP", "E\u00A3", "Egyptian Pound (E\u00A3)"},
    {"VND", "\u20AB",  "Vietnamese Dong (\u20AB)"},
    {"NGN", "\u20A6",  "Nigerian Naira (\u20A6)"},
    {"PKR", "Rs",      "Pakistani Rupee (Rs)"},
    {"BDT", "\u09F3",  "Bangladeshi Taka (\u09F3)"},
    {"UAH", "\u20B4",  "Ukrainian Hryvnia (\u20B4)"},
    {"PEN", "S/.",     "Peruvian Sol (S/.)"},
    {"RON", "lei",     "Romanian Leu (lei)"},
    {"HUF", "Ft",      "Hungarian Forint (Ft)"},
    {"KES", "KSh",     "Kenyan Shilling (KSh)"},
    {"QAR", "QR",      "Qatari Riyal (QR)"},
    {"KWD", "KD",      "Kuwaiti Dinar (KD)"},
    {"BHD", "BD",      "Bahraini Dinar (BD)"},
    {"OMR", "OMR",     "Omani Rial (OMR)"},
    {"LKR", "Rs",      "Sri Lankan Rupee (Rs)"},
    {"MMK", "K",       "Myanmar Kyat (K)"},
};

const QList<CurrencyDef>& NumberFormat::currencies() { return s_currencies; }

QString NumberFormat::getCurrencySymbol(const QString& code) {
    for (const auto& c : s_currencies) if (c.code == code) return c.symbol;
    return "$";
}

NumberFormatType NumberFormat::typeFromString(const QString& str) {
    QString lower = str.toLower();
    if (lower == "number") return NumberFormatType::Number;
    if (lower == "currency") return NumberFormatType::Currency;
    if (lower == "accounting") return NumberFormatType::Accounting;
    if (lower == "percentage") return NumberFormatType::Percentage;
    if (lower == "date") return NumberFormatType::Date;
    if (lower == "time") return NumberFormatType::Time;
    if (lower == "fraction") return NumberFormatType::Fraction;
    if (lower == "scientific") return NumberFormatType::Scientific;
    if (lower == "text") return NumberFormatType::Text;
    if (lower == "custom") return NumberFormatType::Custom;
    if (lower == "special") return NumberFormatType::Special;
    return NumberFormatType::General;
}

QString NumberFormat::typeToString(NumberFormatType type) {
    switch (type) {
        case NumberFormatType::Number: return "Number";
        case NumberFormatType::Currency: return "Currency";
        case NumberFormatType::Accounting: return "Accounting";
        case NumberFormatType::Percentage: return "Percentage";
        case NumberFormatType::Date: return "Date";
        case NumberFormatType::Time: return "Time";
        case NumberFormatType::Fraction: return "Fraction";
        case NumberFormatType::Scientific: return "Scientific";
        case NumberFormatType::Text: return "Text";
        case NumberFormatType::Custom: return "Custom";
        case NumberFormatType::Special: return "Special";
        default: return "General";
    }
}

// ============================================================================
// Helpers
// ============================================================================
static QString formatNumber(double num, int decimals, bool useThousands) {
    QLocale locale(QLocale::English, QLocale::UnitedStates);
    if (useThousands) return locale.toString(num, 'f', decimals);
    return QString::number(num, 'f', decimals);
}

static QDate parseDate(const QString& value) {
    QDate d = QDate::fromString(value, Qt::ISODate);
    if (d.isValid()) return d;
    d = QDate::fromString(value, "MM/dd/yyyy");
    if (d.isValid()) return d;
    d = QDate::fromString(value, "dd/MM/yyyy");
    if (d.isValid()) return d;
    bool ok;
    double serial = value.toDouble(&ok);
    if (ok && serial > 0 && serial < 3000000) {
        return QDate(1899, 12, 30).addDays(static_cast<int>(serial));
    }
    return QDate();
}

// ============================================================================
// Main format function
// ============================================================================
QString NumberFormat::format(const QString& value, const NumberFormatOptions& options) {
    return formatFull(value, options).text;
}

FormatResult NumberFormat::formatFull(const QString& value, const NumberFormatOptions& options) {
    FormatResult result;
    result.text = value;

    if (value.isEmpty() || options.type == NumberFormatType::General ||
        options.type == NumberFormatType::Text) {
        return result;
    }

    bool ok;
    double num = value.toDouble(&ok);

    switch (options.type) {
        case NumberFormatType::Number: {
            if (!ok) return result;
            if (num < 0) {
                double absNum = std::abs(num);
                QString formatted = formatNumber(absNum, options.decimalPlaces, options.useThousandsSeparator);
                switch (options.negativeStyle) {
                    case NegativeStyle::Minus: result.text = "-" + formatted; break;
                    case NegativeStyle::Red: result.text = formatted; result.color = QColor(Qt::red); result.hasColor = true; break;
                    case NegativeStyle::Parentheses: result.text = "(" + formatted + ")"; break;
                    case NegativeStyle::RedParentheses: result.text = "(" + formatted + ")"; result.color = QColor(Qt::red); result.hasColor = true; break;
                }
            } else {
                result.text = formatNumber(num, options.decimalPlaces, options.useThousandsSeparator);
            }
            return result;
        }

        case NumberFormatType::Currency: {
            if (!ok) return result;
            QString symbol = getCurrencySymbol(options.currencyCode);
            QString formatted = formatNumber(std::abs(num), options.decimalPlaces, true);
            if (num < 0) {
                switch (options.negativeStyle) {
                    case NegativeStyle::Minus: result.text = "-" + symbol + formatted; break;
                    case NegativeStyle::Red: result.text = symbol + formatted; result.color = QColor(Qt::red); result.hasColor = true; break;
                    case NegativeStyle::Parentheses: result.text = "(" + symbol + formatted + ")"; break;
                    case NegativeStyle::RedParentheses: result.text = "(" + symbol + formatted + ")"; result.color = QColor(Qt::red); result.hasColor = true; break;
                }
            } else {
                result.text = symbol + formatted;
            }
            return result;
        }

        case NumberFormatType::Accounting: {
            if (!ok) return result;
            QString symbol = getCurrencySymbol(options.currencyCode);
            if (num == 0) {
                result.text = symbol + " -"; // Dash for zero
            } else {
                QString formatted = formatNumber(std::abs(num), options.decimalPlaces, true);
                result.text = (num < 0) ? "(" + symbol + formatted + ")" : " " + symbol + formatted + " ";
            }
            return result;
        }

        case NumberFormatType::Percentage: {
            if (!ok) return result;
            result.text = formatNumber(num * 100.0, options.decimalPlaces, false) + "%";
            return result;
        }

        case NumberFormatType::Date: {
            QDate date = parseDate(value);
            if (!date.isValid()) return result;
            QString fid = options.dateFormatId;
            if (fid == "d/M/yy") result.text = date.toString("d/M/yy");
            else if (fid == "d MMM, yyyy") result.text = date.toString("d MMM, yyyy");
            else if (fid == "d MMMM, yyyy") result.text = date.toString("d MMMM, yyyy");
            else if (fid == "EEEE, d MMMM, yyyy") result.text = date.toString("dddd, d MMMM, yyyy");
            else if (fid == "dd/MM/yyyy") result.text = date.toString("dd/MM/yyyy");
            else if (fid == "MM/dd/yyyy" || fid == "mm/dd/yyyy") result.text = date.toString("MM/dd/yyyy");
            else if (fid == "yyyy/MM/dd") result.text = date.toString("yyyy/MM/dd");
            else if (fid == "yyyy-MM-dd" || fid == "yyyy-mm-dd") result.text = date.toString("yyyy-MM-dd");
            else if (fid == "mmm d, yyyy") result.text = date.toString("MMM d, yyyy");
            else if (fid == "mmmm d, yyyy") result.text = date.toString("MMMM d, yyyy");
            else if (fid == "d-mmm-yy") result.text = date.toString("d-MMM-yy");
            else if (fid == "mm/dd") result.text = date.toString("MM/dd");
            else result.text = date.toString("MM/dd/yyyy");
            return result;
        }

        case NumberFormatType::Time: {
            if (!ok) return result;
            double fraction = num - std::floor(num);
            int totalSecs = static_cast<int>(std::round(fraction * 86400));
            int h = totalSecs / 3600;
            int m = (totalSecs % 3600) / 60;
            int s = totalSecs % 60;
            QString fid = options.dateFormatId;
            if (fid == "h:mm") result.text = QString("%1:%2").arg(h).arg(m, 2, 10, QChar('0'));
            else if (fid == "h:mm:ss") result.text = QString("%1:%2:%3").arg(h).arg(m, 2, 10, QChar('0')).arg(s, 2, 10, QChar('0'));
            else if (fid == "h:mm AM/PM") {
                QString ampm = h >= 12 ? "PM" : "AM";
                int h12 = h % 12; if (h12 == 0) h12 = 12;
                result.text = QString("%1:%2 %3").arg(h12).arg(m, 2, 10, QChar('0')).arg(ampm);
            } else if (fid == "h:mm:ss AM/PM") {
                QString ampm = h >= 12 ? "PM" : "AM";
                int h12 = h % 12; if (h12 == 0) h12 = 12;
                result.text = QString("%1:%2:%3 %4").arg(h12).arg(m, 2, 10, QChar('0')).arg(s, 2, 10, QChar('0')).arg(ampm);
            } else if (fid == "[h]:mm:ss") {
                // Elapsed time (hours can exceed 24)
                int totalH = static_cast<int>(num * 24);
                int totalM = static_cast<int>(num * 1440) % 60;
                int totalS = static_cast<int>(num * 86400) % 60;
                result.text = QString("%1:%2:%3").arg(totalH).arg(totalM, 2, 10, QChar('0')).arg(totalS, 2, 10, QChar('0'));
            } else {
                // Default: h:mm:ss AM/PM
                QString ampm = h >= 12 ? "PM" : "AM";
                int h12 = h % 12; if (h12 == 0) h12 = 12;
                result.text = QString("%1:%2:%3 %4").arg(h12).arg(m, 2, 10, QChar('0')).arg(s, 2, 10, QChar('0')).arg(ampm);
            }
            return result;
        }

        case NumberFormatType::Scientific: {
            if (!ok) return result;
            result.text = QString::number(num, 'E', options.decimalPlaces);
            return result;
        }

        case NumberFormatType::Fraction: {
            if (!ok) return result;
            int whole = static_cast<int>(num);
            double frac = std::abs(num - whole);
            if (frac < 0.0001) { result.text = QString::number(whole); return result; }
            int bestNum = 0, bestDen = 1;
            double bestErr = 1.0;
            for (int den = 1; den <= 100; ++den) {
                int num_ = static_cast<int>(std::round(frac * den));
                double err = std::abs(frac - static_cast<double>(num_) / den);
                if (err < bestErr) { bestErr = err; bestNum = num_; bestDen = den; }
                if (bestErr < 0.0001) break;
            }
            if (bestNum == 0) { result.text = QString::number(whole); return result; }
            if (whole == 0) result.text = QString("%1/%2").arg(num < 0 ? -bestNum : bestNum).arg(bestDen);
            else result.text = QString("%1 %2/%3").arg(whole).arg(bestNum).arg(bestDen);
            return result;
        }

        case NumberFormatType::Special: {
            if (!ok) return result;
            long long n = static_cast<long long>(num);
            if (options.specialType == "zipcode") {
                result.text = QString("%1").arg(n, 5, 10, QChar('0'));
            } else if (options.specialType == "zipcode4") {
                result.text = QString("%1-%2").arg(n / 10000, 5, 10, QChar('0')).arg(n % 10000, 4, 10, QChar('0'));
            } else if (options.specialType == "phone") {
                result.text = QString("(%1) %2-%3")
                    .arg(n / 10000000, 3, 10, QChar('0'))
                    .arg((n / 10000) % 1000, 3, 10, QChar('0'))
                    .arg(n % 10000, 4, 10, QChar('0'));
            } else if (options.specialType == "ssn") {
                result.text = QString("%1-%2-%3")
                    .arg(n / 10000000, 3, 10, QChar('0'))
                    .arg((n / 10000) % 100, 2, 10, QChar('0'))
                    .arg(n % 10000, 4, 10, QChar('0'));
            } else {
                result.text = value;
            }
            return result;
        }

        case NumberFormatType::Custom:
            return applyCustomFormatFull(value, options.customFormat);

        default:
            return result;
    }
}

// ============================================================================
// Custom format code engine (Excel-compatible)
// ============================================================================
FormatResult NumberFormat::applyCustomFormatFull(const QString& value, const QString& formatStr) {
    FormatResult result;
    result.text = value;
    if (formatStr.isEmpty()) return result;

    bool ok;
    double num = value.toDouble(&ok);

    // Split format for positive;negative;zero;text (up to 4 sections)
    QStringList sections = formatStr.split(';');

    // Select the appropriate section
    QString fmt;
    bool isText = !ok;
    if (isText) {
        fmt = sections.size() > 3 ? sections[3] : (sections.size() > 0 ? sections[0] : formatStr);
        // Replace @ with cell text
        if (fmt.contains('@')) {
            result.text = fmt;
            result.text.replace("@", value);
            return result;
        }
        result.text = value;
        return result;
    }

    if (num > 0 || sections.size() == 1) fmt = sections[0];
    else if (num < 0 && sections.size() > 1) fmt = sections[1];
    else if (num == 0 && sections.size() > 2) fmt = sections[2];
    else fmt = sections[0];

    // Extract color code [Red], [Blue], etc.
    static QRegularExpression colorRe("\\[(Red|Blue|Green|Black|White|Magenta|Cyan|Yellow)\\]",
                                       QRegularExpression::CaseInsensitiveOption);
    auto colorMatch = colorRe.match(fmt);
    if (colorMatch.hasMatch()) {
        QString colorName = colorMatch.captured(1).toLower();
        if (colorName == "red") result.color = QColor(Qt::red);
        else if (colorName == "blue") result.color = QColor(Qt::blue);
        else if (colorName == "green") result.color = QColor(Qt::darkGreen);
        else if (colorName == "black") result.color = QColor(Qt::black);
        else if (colorName == "magenta") result.color = QColor(Qt::magenta);
        else if (colorName == "cyan") result.color = QColor(Qt::cyan);
        else if (colorName == "yellow") result.color = QColor(Qt::darkYellow);
        result.hasColor = true;
        fmt.replace(colorRe, "");
    }

    // Extract conditions [>1000], [<=500], etc.
    static QRegularExpression condRe("\\[([<>=!]+)(\\d+\\.?\\d*)\\]");
    auto condMatch = condRe.match(fmt);
    if (condMatch.hasMatch()) {
        QString op = condMatch.captured(1);
        double threshold = condMatch.captured(2).toDouble();
        bool condMet = false;
        if (op == ">") condMet = num > threshold;
        else if (op == "<") condMet = num < threshold;
        else if (op == ">=") condMet = num >= threshold;
        else if (op == "<=") condMet = num <= threshold;
        else if (op == "=" || op == "==") condMet = num == threshold;
        else if (op == "<>" || op == "!=") condMet = num != threshold;
        if (!condMet && sections.size() > 1) {
            // Use second section
            return applyCustomFormatFull(value, sections[1]);
        }
        fmt.replace(condRe, "");
    }

    // Handle @ (text placeholder)
    if (fmt.contains('@')) {
        result.text = fmt;
        result.text.replace("@", value);
        return result;
    }

    // Handle date/time codes in custom format
    if (fmt.contains("yyyy") || fmt.contains("mm") || fmt.contains("dd") ||
        fmt.contains("hh") || fmt.contains("ss") || fmt.contains("AM/PM")) {
        QDate date = parseDate(value);
        if (date.isValid()) {
            QString out = fmt;
            out.replace("yyyy", date.toString("yyyy"));
            out.replace("yy", date.toString("yy"));
            out.replace("mmmm", date.toString("MMMM"));
            out.replace("mmm", date.toString("MMM"));
            out.replace("mm", date.toString("MM"));
            out.replace("dddd", date.toString("dddd"));
            out.replace("ddd", date.toString("ddd"));
            out.replace("dd", date.toString("dd"));
            out.replace("d", date.toString("d"));
            // Time portion
            if (ok) {
                double frac = num - std::floor(num);
                int totalSecs = static_cast<int>(std::round(frac * 86400));
                int h = totalSecs / 3600, m = (totalSecs % 3600) / 60, s = totalSecs % 60;
                bool hasAMPM = out.contains("AM/PM");
                if (hasAMPM) {
                    out.replace("AM/PM", h >= 12 ? "PM" : "AM");
                    h = h % 12; if (h == 0) h = 12;
                }
                out.replace("hh", QString("%1").arg(h, 2, 10, QChar('0')));
                out.replace("h", QString::number(h));
                out.replace("ss", QString("%1").arg(s, 2, 10, QChar('0')));
                // mm after hh means minutes, not months
                out.replace("mm", QString("%1").arg(m, 2, 10, QChar('0')));
            }
            result.text = out;
            return result;
        }
    }

    // Numeric formatting
    bool isPercent = fmt.contains('%');
    double val = isPercent ? num * 100.0 : num;
    if (num < 0 && sections.size() > 1) val = std::abs(val); // Negative section handles sign

    // Handle scaling (trailing commas divide by 1000)
    int scalingCommas = 0;
    QString fmtClean = fmt;
    while (fmtClean.endsWith(',') || fmtClean.endsWith(",\"") || fmtClean.endsWith(",%")) {
        if (fmtClean.endsWith(',')) { scalingCommas++; fmtClean.chop(1); }
        else break;
    }
    // Also check before % or space
    {
        QRegularExpression trailingComma(",+(?=[%\"\\s]|$)");
        auto m = trailingComma.match(fmtClean);
        if (m.hasMatch()) {
            scalingCommas += m.capturedLength();
            fmtClean.replace(trailingComma, "");
        }
    }
    for (int i = 0; i < scalingCommas; ++i) val /= 1000.0;

    // Count decimal places from format
    int dotIdx = fmtClean.indexOf('.');
    int decimals = 0;
    if (dotIdx >= 0) {
        for (int i = dotIdx + 1; i < fmtClean.length(); ++i) {
            if (fmtClean[i] == '0' || fmtClean[i] == '#' || fmtClean[i] == '?') decimals++;
            else break;
        }
    }

    bool useComma = fmtClean.contains(',');

    // Handle _ (skip width of next char) — just remove it
    fmtClean.replace(QRegularExpression("_."), " ");

    // Handle * (fill character) — just remove it
    fmtClean.replace(QRegularExpression("\\*."), "");

    // Extract prefix and suffix
    QString prefix, suffix;
    // Find prefix: everything before first digit placeholder
    int firstDigit = -1;
    for (int i = 0; i < fmtClean.length(); ++i) {
        if (fmtClean[i] == '0' || fmtClean[i] == '#' || fmtClean[i] == '?') {
            firstDigit = i; break;
        }
    }
    if (firstDigit > 0) {
        prefix = fmtClean.left(firstDigit);
        prefix.remove('"');
        prefix.remove('\\');
    }
    // Find suffix: everything after last digit placeholder
    int lastDigit = -1;
    for (int i = fmtClean.length() - 1; i >= 0; --i) {
        if (fmtClean[i] == '0' || fmtClean[i] == '#' || fmtClean[i] == '?') {
            lastDigit = i; break;
        }
    }
    if (lastDigit >= 0 && lastDigit < fmtClean.length() - 1) {
        suffix = fmtClean.mid(lastDigit + 1);
        suffix.remove('"');
        suffix.remove('\\');
    }

    if (isPercent && !suffix.contains('%')) suffix += "%";

    QString formatted = formatNumber(std::abs(val), decimals, useComma);

    // Add sign for negative in first section
    QString sign;
    if (num < 0 && sections.size() <= 1) sign = "-";

    result.text = sign + prefix + formatted + suffix;
    return result;
}

QString NumberFormat::applyCustomFormat(const QString& value, const QString& formatStr) {
    return applyCustomFormatFull(value, formatStr).text;
}
