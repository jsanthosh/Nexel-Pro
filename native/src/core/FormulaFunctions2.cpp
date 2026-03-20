// ============================================================================
// FormulaFunctions2.cpp — Extended function library batch 2
// ============================================================================
// Additional P1 functions from EXCEL_FEATURES.md Section 13.
//

#include "FormulaEngine.h"
#include "Spreadsheet.h"
#include <cmath>
#include <algorithm>
#include <numeric>
#include <random>
#include <set>
#include <map>
#include <QDateTime>
#include <QRegularExpression>

static const QDate s_epoch(1899, 12, 30);
static double dateSerial(const QDate& d) { return s_epoch.daysTo(d); }
static QDate serialDate(double s) { return s_epoch.addDays(static_cast<int>(s)); }

// ============================================================================
// MATH & TRIG — Batch 2
// ============================================================================

// SINH, COSH, TANH
QVariant FormulaEngine::funcSINH(const std::vector<QVariant>& args) {
    return args.empty() ? QVariant("#VALUE!") : QVariant(std::sinh(toNumber(args[0])));
}
QVariant FormulaEngine::funcCOSH(const std::vector<QVariant>& args) {
    return args.empty() ? QVariant("#VALUE!") : QVariant(std::cosh(toNumber(args[0])));
}
QVariant FormulaEngine::funcTANH(const std::vector<QVariant>& args) {
    return args.empty() ? QVariant("#VALUE!") : QVariant(std::tanh(toNumber(args[0])));
}

// FACTDOUBLE(n) — n!! = n*(n-2)*(n-4)*...
QVariant FormulaEngine::funcFACTDOUBLE(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    int n = static_cast<int>(toNumber(args[0]));
    if (n < -1) return QVariant("#NUM!");
    if (n <= 0) return 1.0;
    double result = 1;
    for (int i = n; i > 0; i -= 2) result *= i;
    return result;
}

// MULTINOMIAL(n1, n2, ...) = (n1+n2+...)! / (n1! * n2! * ...)
QVariant FormulaEngine::funcMULTINOMIAL(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args);
    int sum = 0;
    double denom = 1;
    for (const auto& a : flat) {
        int n = static_cast<int>(toNumber(a));
        if (n < 0) return QVariant("#NUM!");
        sum += n;
        double f = 1; for (int i = 2; i <= n; ++i) f *= i;
        denom *= f;
    }
    double num = 1; for (int i = 2; i <= sum; ++i) num *= i;
    return num / denom;
}

// BASE(number, radix, [min_length])
QVariant FormulaEngine::funcBASE(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    long long n = static_cast<long long>(toNumber(args[0]));
    int radix = static_cast<int>(toNumber(args[1]));
    int minLen = args.size() > 2 ? static_cast<int>(toNumber(args[2])) : 0;
    if (radix < 2 || radix > 36) return QVariant("#VALUE!");
    QString result = QString::number(n, radix).toUpper();
    while (result.length() < minLen) result = "0" + result;
    return result;
}

// DECIMAL(text, radix)
QVariant FormulaEngine::funcDECIMAL(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    QString text = toString(args[0]);
    int radix = static_cast<int>(toNumber(args[1]));
    if (radix < 2 || radix > 36) return QVariant("#VALUE!");
    bool ok;
    long long result = text.toLongLong(&ok, radix);
    return ok ? QVariant(static_cast<double>(result)) : QVariant("#VALUE!");
}

// ROMAN(number)
QVariant FormulaEngine::funcROMAN(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    int n = static_cast<int>(toNumber(args[0]));
    if (n < 0 || n > 3999) return QVariant("#VALUE!");
    if (n == 0) return QString("");
    const int vals[] = {1000,900,500,400,100,90,50,40,10,9,5,4,1};
    const char* syms[] = {"M","CM","D","CD","C","XC","L","XL","X","IX","V","IV","I"};
    QString result;
    for (int i = 0; i < 13; ++i) {
        while (n >= vals[i]) { result += syms[i]; n -= vals[i]; }
    }
    return result;
}

// ARABIC(roman_text)
QVariant FormulaEngine::funcARABIC(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    QString s = toString(args[0]).toUpper();
    auto romanVal = [](QChar c) -> int {
        switch (c.unicode()) {
            case 'I': return 1; case 'V': return 5; case 'X': return 10;
            case 'L': return 50; case 'C': return 100; case 'D': return 500; case 'M': return 1000;
            default: return 0;
        }
    };
    int result = 0;
    for (int i = 0; i < s.length(); ++i) {
        int cur = romanVal(s[i]);
        int next = (i + 1 < s.length()) ? romanVal(s[i+1]) : 0;
        result += (cur < next) ? -cur : cur;
    }
    return result;
}

// SUBTOTAL(function_num, ref1, ...)
QVariant FormulaEngine::funcSUBTOTAL(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    int funcNum = static_cast<int>(toNumber(args[0]));
    std::vector<QVariant> dataArgs(args.begin() + 1, args.end());

    // Function numbers: 1-11 or 101-111 (101+ ignore hidden rows)
    int fn = (funcNum > 100) ? funcNum - 100 : funcNum;
    switch (fn) {
        case 1: return funcAVERAGE(dataArgs);
        case 2: return funcCOUNT(dataArgs);
        case 3: return funcCOUNTA(dataArgs);
        case 4: return funcMAX(dataArgs);
        case 5: return funcMIN(dataArgs);
        case 6: return funcPRODUCT(dataArgs);
        case 7: return funcSTDEV(dataArgs);
        case 8: return funcSTDEV(dataArgs); // STDEVP approximation
        case 9: return funcSUM(dataArgs);
        case 10: return funcVAR(dataArgs);
        case 11: return funcVAR(dataArgs); // VARP approximation
        default: return QVariant("#VALUE!");
    }
}

// COMBINA(n, k) = COMBIN(n+k-1, k)
QVariant FormulaEngine::funcCOMBINA(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    int n = static_cast<int>(toNumber(args[0]));
    int k = static_cast<int>(toNumber(args[1]));
    if (n < 0 || k < 0) return QVariant("#NUM!");
    std::vector<QVariant> newArgs = {QVariant(static_cast<double>(n + k - 1)), QVariant(static_cast<double>(k))};
    return funcCOMBIN(newArgs);
}

// SERIESSUM(x, n, m, coefficients)
QVariant FormulaEngine::funcSERIESSUM(const std::vector<QVariant>& args) {
    if (args.size() < 4) return QVariant("#VALUE!");
    double x = toNumber(args[0]);
    double n = toNumber(args[1]);
    double m = toNumber(args[2]);
    auto coeffs = flattenArgs({args[3]});
    double sum = 0;
    for (size_t i = 0; i < coeffs.size(); ++i) {
        sum += toNumber(coeffs[i]) * std::pow(x, n + i * m);
    }
    return sum;
}

// ============================================================================
// LOOKUP & REFERENCE — Batch 2
// ============================================================================

// XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])
QVariant FormulaEngine::funcXMATCH(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    QVariant lookup = args[0];
    auto array = flattenArgs({args[1]});
    int matchMode = args.size() > 2 ? static_cast<int>(toNumber(args[2])) : 0;
    // 0=exact, -1=exact or next smaller, 1=exact or next larger, 2=wildcard

    QString lookupStr = toString(lookup);
    double lookupNum = toNumber(lookup);
    bool isNum = (lookup.typeId() == QMetaType::Double || lookup.typeId() == QMetaType::Int);

    for (int i = 0; i < static_cast<int>(array.size()); ++i) {
        if (matchMode == 0 || matchMode == 2) {
            if (matchMode == 2) {
                // Wildcard match
                QString pattern = toString(array[i]);
                pattern.replace("*", ".*").replace("?", ".");
                QRegularExpression re("^" + pattern + "$", QRegularExpression::CaseInsensitiveOption);
                if (re.match(lookupStr).hasMatch()) return i + 1;
            } else {
                if (isNum && toNumber(array[i]) == lookupNum) return i + 1;
                if (!isNum && toString(array[i]).compare(lookupStr, Qt::CaseInsensitive) == 0) return i + 1;
            }
        }
    }

    // For approximate modes, find closest
    if (matchMode == -1 || matchMode == 1) {
        int bestIdx = -1;
        double bestDiff = std::numeric_limits<double>::max();
        for (int i = 0; i < static_cast<int>(array.size()); ++i) {
            double v = toNumber(array[i]);
            double diff = lookupNum - v;
            if (matchMode == -1 && diff >= 0 && diff < bestDiff) { bestDiff = diff; bestIdx = i; }
            if (matchMode == 1 && diff <= 0 && std::abs(diff) < bestDiff) { bestDiff = std::abs(diff); bestIdx = i; }
        }
        if (bestIdx >= 0) return bestIdx + 1;
    }

    return QVariant("#N/A");
}

// SORTBY(array, by_array1, [sort_order1], ...)
QVariant FormulaEngine::funcSORTBY(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    auto data = flattenArgs({args[0]});
    auto keys = flattenArgs({args[1]});
    int order = args.size() > 2 ? static_cast<int>(toNumber(args[2])) : 1;
    bool ascending = (order == 1);

    int n = std::min(data.size(), keys.size());
    std::vector<int> indices(n);
    std::iota(indices.begin(), indices.end(), 0);

    std::sort(indices.begin(), indices.end(), [&](int a, int b) {
        double ka = toNumber(keys[a]), kb = toNumber(keys[b]);
        return ascending ? ka < kb : ka > kb;
    });

    std::vector<QVariant> result;
    for (int i : indices) result.push_back(data[i]);
    return QVariant::fromValue(result);
}

// VSTACK(array1, array2, ...) — stack arrays vertically
QVariant FormulaEngine::funcVSTACK(const std::vector<QVariant>& args) {
    std::vector<QVariant> result;
    for (const auto& arg : args) {
        if (arg.canConvert<std::vector<QVariant>>()) {
            auto v = arg.value<std::vector<QVariant>>();
            result.insert(result.end(), v.begin(), v.end());
        } else {
            result.push_back(arg);
        }
    }
    return QVariant::fromValue(result);
}

// HSTACK(array1, array2, ...) — stack arrays horizontally (simplified: concat)
QVariant FormulaEngine::funcHSTACK(const std::vector<QVariant>& args) {
    // Simplified: treat as row concatenation
    return funcVSTACK(args);
}

// TAKE(array, rows, [columns]) — take first/last N rows
QVariant FormulaEngine::funcTAKE(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    auto data = flattenArgs({args[0]});
    int rows = static_cast<int>(toNumber(args[1]));
    if (rows == 0 || data.empty()) return QVariant("#VALUE!");

    std::vector<QVariant> result;
    if (rows > 0) {
        int n = std::min(rows, static_cast<int>(data.size()));
        result.assign(data.begin(), data.begin() + n);
    } else {
        int n = std::min(-rows, static_cast<int>(data.size()));
        result.assign(data.end() - n, data.end());
    }
    return QVariant::fromValue(result);
}

// DROP(array, rows, [columns]) — drop first/last N rows
QVariant FormulaEngine::funcDROP(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    auto data = flattenArgs({args[0]});
    int rows = static_cast<int>(toNumber(args[1]));

    std::vector<QVariant> result;
    if (rows > 0 && rows < static_cast<int>(data.size())) {
        result.assign(data.begin() + rows, data.end());
    } else if (rows < 0 && -rows < static_cast<int>(data.size())) {
        result.assign(data.begin(), data.end() + rows);
    }
    return QVariant::fromValue(result);
}

// ============================================================================
// TEXT — Batch 2
// ============================================================================

// DOLLAR(number, [decimals])
QVariant FormulaEngine::funcDOLLAR(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double n = toNumber(args[0]);
    int dec = args.size() > 1 ? static_cast<int>(toNumber(args[1])) : 2;
    QString formatted = QString::number(std::abs(n), 'f', dec);
    // Add thousands separator
    int dotPos = formatted.indexOf('.');
    if (dotPos < 0) dotPos = formatted.length();
    for (int i = dotPos - 3; i > 0; i -= 3) formatted.insert(i, ',');
    return (n < 0 ? "-$" : "$") + formatted;
}

// CONCAT(text1, text2, ...) — already exists as funcCONCAT
// TEXTJOIN — already exists

// LENB(text) — byte length (simplified: same as LEN for non-DBCS)
QVariant FormulaEngine::funcLENB(const std::vector<QVariant>& args) {
    return funcLEN(args);
}

// FINDB — same as FIND for non-DBCS
QVariant FormulaEngine::funcFINDB(const std::vector<QVariant>& args) {
    return funcFIND(args);
}

// SEARCHB — same as SEARCH for non-DBCS
QVariant FormulaEngine::funcSEARCHB(const std::vector<QVariant>& args) {
    return funcSEARCH(args);
}

// VALUETOTEXT(value, [format])
QVariant FormulaEngine::funcVALUETOTEXT(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    return toString(args[0]);
}

// ============================================================================
// STATISTICAL — Batch 2
// ============================================================================

// STDEV.P (population standard deviation)
QVariant FormulaEngine::funcSTDEVP(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args);
    if (flat.empty()) return QVariant("#DIV/0!");
    double sum = 0, sum2 = 0; int n = 0;
    for (const auto& a : flat) {
        if (!a.isNull() && a.isValid()) {
            double v = toNumber(a); sum += v; sum2 += v * v; n++;
        }
    }
    if (n == 0) return QVariant("#DIV/0!");
    double mean = sum / n;
    return std::sqrt(sum2 / n - mean * mean);
}

// VAR.P (population variance)
QVariant FormulaEngine::funcVARP(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args);
    if (flat.empty()) return QVariant("#DIV/0!");
    double sum = 0, sum2 = 0; int n = 0;
    for (const auto& a : flat) {
        if (!a.isNull() && a.isValid()) {
            double v = toNumber(a); sum += v; sum2 += v * v; n++;
        }
    }
    if (n == 0) return QVariant("#DIV/0!");
    double mean = sum / n;
    return sum2 / n - mean * mean;
}

// PERCENTILE.INC(array, k)
QVariant FormulaEngine::funcPERCENTILE_INC(const std::vector<QVariant>& args) {
    return funcPERCENTILE(args); // Same as PERCENTILE
}

// PERCENTILE.EXC(array, k)
QVariant FormulaEngine::funcPERCENTILE_EXC(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    auto flat = flattenArgs({args[0]});
    double k = toNumber(args[1]);
    std::vector<double> sorted;
    for (const auto& a : flat) if (!a.isNull() && a.isValid()) sorted.push_back(toNumber(a));
    std::sort(sorted.begin(), sorted.end());
    int n = sorted.size();
    if (n == 0 || k <= 0 || k >= 1) return QVariant("#NUM!");
    double idx = k * (n + 1) - 1;
    int lo = static_cast<int>(std::floor(idx));
    double frac = idx - lo;
    if (lo < 0) return sorted[0];
    if (lo >= n - 1) return sorted[n - 1];
    return sorted[lo] + frac * (sorted[lo + 1] - sorted[lo]);
}

// QUARTILE.INC(array, quart)
QVariant FormulaEngine::funcQUARTILE_INC(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    int q = static_cast<int>(toNumber(args[1]));
    if (q < 0 || q > 4) return QVariant("#NUM!");
    std::vector<QVariant> newArgs = {args[0], QVariant(q * 0.25)};
    return funcPERCENTILE(newArgs);
}

// RANK.EQ(number, ref, [order])
QVariant FormulaEngine::funcRANK_EQ(const std::vector<QVariant>& args) {
    return funcRANK(args); // Same behavior
}

// RANK.AVG(number, ref, [order])
QVariant FormulaEngine::funcRANK_AVG(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    double num = toNumber(args[0]);
    auto flat = flattenArgs({args[1]});
    int order = args.size() > 2 ? static_cast<int>(toNumber(args[2])) : 0;

    std::vector<double> values;
    for (const auto& a : flat) if (!a.isNull() && a.isValid()) values.push_back(toNumber(a));

    int rank = 1, count = 0;
    for (double v : values) {
        if (order == 0) { if (v > num) rank++; }
        else { if (v < num) rank++; }
        if (v == num) count++;
    }
    return rank + (count - 1) / 2.0; // Average rank for ties
}

// GEOMEAN(number1, ...)
QVariant FormulaEngine::funcGEOMEAN(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args);
    double logSum = 0; int n = 0;
    for (const auto& a : flat) {
        if (!a.isNull() && a.isValid()) {
            double v = toNumber(a);
            if (v <= 0) return QVariant("#NUM!");
            logSum += std::log(v); n++;
        }
    }
    if (n == 0) return QVariant("#NUM!");
    return std::exp(logSum / n);
}

// HARMEAN(number1, ...)
QVariant FormulaEngine::funcHARMEAN(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args);
    double recipSum = 0; int n = 0;
    for (const auto& a : flat) {
        if (!a.isNull() && a.isValid()) {
            double v = toNumber(a);
            if (v <= 0) return QVariant("#NUM!");
            recipSum += 1.0 / v; n++;
        }
    }
    if (n == 0 || recipSum == 0) return QVariant("#NUM!");
    return n / recipSum;
}

// TRIMMEAN(array, percent)
QVariant FormulaEngine::funcTRIMMEAN(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    auto flat = flattenArgs({args[0]});
    double pct = toNumber(args[1]);
    if (pct < 0 || pct >= 1) return QVariant("#NUM!");

    std::vector<double> sorted;
    for (const auto& a : flat) if (!a.isNull() && a.isValid()) sorted.push_back(toNumber(a));
    std::sort(sorted.begin(), sorted.end());

    int n = sorted.size();
    int trim = static_cast<int>(std::floor(n * pct / 2));
    double sum = 0; int count = 0;
    for (int i = trim; i < n - trim; ++i) { sum += sorted[i]; count++; }
    return count > 0 ? sum / count : QVariant("#DIV/0!");
}

// DEVSQ(number1, ...) — sum of squared deviations from mean
QVariant FormulaEngine::funcDEVSQ(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args);
    double sum = 0; int n = 0;
    for (const auto& a : flat) {
        if (!a.isNull() && a.isValid()) { sum += toNumber(a); n++; }
    }
    if (n == 0) return QVariant("#NUM!");
    double mean = sum / n;
    double devsq = 0;
    for (const auto& a : flat) {
        if (!a.isNull() && a.isValid()) { double d = toNumber(a) - mean; devsq += d * d; }
    }
    return devsq;
}

// AVEDEV(number1, ...) — average absolute deviation
QVariant FormulaEngine::funcAVEDEV(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args);
    double sum = 0; int n = 0;
    for (const auto& a : flat) {
        if (!a.isNull() && a.isValid()) { sum += toNumber(a); n++; }
    }
    if (n == 0) return QVariant("#NUM!");
    double mean = sum / n;
    double absdev = 0;
    for (const auto& a : flat) {
        if (!a.isNull() && a.isValid()) absdev += std::abs(toNumber(a) - mean);
    }
    return absdev / n;
}

// RSQ(known_y, known_x) — R-squared
QVariant FormulaEngine::funcRSQ(const std::vector<QVariant>& args) {
    QVariant corr = funcCORREL(args);
    if (corr.toString().startsWith("#")) return corr;
    double r = corr.toDouble();
    return r * r;
}

// FREQUENCY(data_array, bins_array) — returns frequency distribution
QVariant FormulaEngine::funcFREQUENCY(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    auto data = flattenArgs({args[0]});
    auto bins = flattenArgs({args[1]});

    std::vector<double> binVals;
    for (const auto& b : bins) binVals.push_back(toNumber(b));
    std::sort(binVals.begin(), binVals.end());

    std::vector<int> freq(binVals.size() + 1, 0);
    for (const auto& d : data) {
        double v = toNumber(d);
        int idx = static_cast<int>(std::upper_bound(binVals.begin(), binVals.end(), v) - binVals.begin());
        freq[idx]++;
    }

    std::vector<QVariant> result;
    for (int f : freq) result.push_back(QVariant(f));
    return QVariant::fromValue(result);
}

// ============================================================================
// FINANCIAL — Batch 2
// ============================================================================

// RATE(nper, pmt, pv, [fv], [type], [guess])
QVariant FormulaEngine::funcRATE(const std::vector<QVariant>& args) {
    if (args.size() < 3) return QVariant("#VALUE!");
    double nper = toNumber(args[0]);
    double pmt = toNumber(args[1]);
    double pv = toNumber(args[2]);
    double fv = args.size() > 3 ? toNumber(args[3]) : 0;
    int type = args.size() > 4 ? static_cast<int>(toNumber(args[4])) : 0;
    double guess = args.size() > 5 ? toNumber(args[5]) : 0.1;

    // Newton-Raphson
    double rate = guess;
    for (int iter = 0; iter < 100; ++iter) {
        double pvif = std::pow(1 + rate, nper);
        double f = pv * pvif + pmt * (1 + rate * type) * (pvif - 1) / rate + fv;
        double df = pv * nper * std::pow(1 + rate, nper - 1)
                   + pmt * (1 + rate * type) * (nper * std::pow(1 + rate, nper - 1) * rate - (pvif - 1)) / (rate * rate)
                   + pmt * type * (pvif - 1) / rate;
        if (std::abs(df) < 1e-15) break;
        double newRate = rate - f / df;
        if (std::abs(newRate - rate) < 1e-10) return newRate;
        rate = newRate;
    }
    return rate;
}

// XNPV(rate, values, dates)
QVariant FormulaEngine::funcXNPV(const std::vector<QVariant>& args) {
    if (args.size() < 3) return QVariant("#VALUE!");
    double rate = toNumber(args[0]);
    auto values = flattenArgs({args[1]});
    auto dates = flattenArgs({args[2]});
    if (values.size() != dates.size() || values.empty()) return QVariant("#VALUE!");

    double d0 = toNumber(dates[0]);
    double npv = 0;
    for (size_t i = 0; i < values.size(); ++i) {
        double years = (toNumber(dates[i]) - d0) / 365.0;
        npv += toNumber(values[i]) / std::pow(1 + rate, years);
    }
    return npv;
}

// XIRR(values, dates, [guess])
QVariant FormulaEngine::funcXIRR(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    auto values = flattenArgs({args[0]});
    auto dates = flattenArgs({args[1]});
    double guess = args.size() > 2 ? toNumber(args[2]) : 0.1;

    if (values.size() != dates.size() || values.empty()) return QVariant("#VALUE!");
    double d0 = toNumber(dates[0]);

    double rate = guess;
    for (int iter = 0; iter < 100; ++iter) {
        double npv = 0, dnpv = 0;
        for (size_t i = 0; i < values.size(); ++i) {
            double years = (toNumber(dates[i]) - d0) / 365.0;
            double pv = toNumber(values[i]) / std::pow(1 + rate, years);
            npv += pv;
            dnpv -= years * pv / (1 + rate);
        }
        if (std::abs(dnpv) < 1e-15) break;
        double newRate = rate - npv / dnpv;
        if (std::abs(newRate - rate) < 1e-10) return newRate;
        rate = newRate;
    }
    return rate;
}

// CUMIPMT(rate, nper, pv, start_period, end_period, type)
QVariant FormulaEngine::funcCUMIPMT(const std::vector<QVariant>& args) {
    if (args.size() < 6) return QVariant("#VALUE!");
    double rate = toNumber(args[0]);
    double nper = toNumber(args[1]);
    double pv = toNumber(args[2]);
    int start = static_cast<int>(toNumber(args[3]));
    int end = static_cast<int>(toNumber(args[4]));
    int type = static_cast<int>(toNumber(args[5]));

    double sum = 0;
    for (int per = start; per <= end; ++per) {
        std::vector<QVariant> ipmtArgs = {QVariant(rate), QVariant(static_cast<double>(per)),
                                           QVariant(nper), QVariant(pv), QVariant(0.0), QVariant(static_cast<double>(type))};
        sum += toNumber(funcIPMT(ipmtArgs));
    }
    return sum;
}

// CUMPRINC(rate, nper, pv, start_period, end_period, type)
QVariant FormulaEngine::funcCUMPRINC(const std::vector<QVariant>& args) {
    if (args.size() < 6) return QVariant("#VALUE!");
    double rate = toNumber(args[0]);
    double nper = toNumber(args[1]);
    double pv = toNumber(args[2]);
    int start = static_cast<int>(toNumber(args[3]));
    int end = static_cast<int>(toNumber(args[4]));
    int type = static_cast<int>(toNumber(args[5]));

    double sum = 0;
    for (int per = start; per <= end; ++per) {
        std::vector<QVariant> ppmtArgs = {QVariant(rate), QVariant(static_cast<double>(per)),
                                           QVariant(nper), QVariant(pv), QVariant(0.0), QVariant(static_cast<double>(type))};
        sum += toNumber(funcPPMT(ppmtArgs));
    }
    return sum;
}

// PDURATION(rate, pv, fv)
QVariant FormulaEngine::funcPDURATION(const std::vector<QVariant>& args) {
    if (args.size() < 3) return QVariant("#VALUE!");
    double rate = toNumber(args[0]);
    double pv = toNumber(args[1]);
    double fv = toNumber(args[2]);
    if (rate <= 0 || pv <= 0 || fv <= 0) return QVariant("#NUM!");
    return (std::log(fv) - std::log(pv)) / std::log(1 + rate);
}

// RRI(nper, pv, fv)
QVariant FormulaEngine::funcRRI(const std::vector<QVariant>& args) {
    if (args.size() < 3) return QVariant("#VALUE!");
    double nper = toNumber(args[0]);
    double pv = toNumber(args[1]);
    double fv = toNumber(args[2]);
    if (nper <= 0 || pv == 0) return QVariant("#NUM!");
    return std::pow(fv / pv, 1.0 / nper) - 1;
}

// FVSCHEDULE(principal, schedule)
QVariant FormulaEngine::funcFVSCHEDULE(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    double pv = toNumber(args[0]);
    auto rates = flattenArgs({args[1]});
    for (const auto& r : rates) pv *= (1 + toNumber(r));
    return pv;
}

// ============================================================================
// DATE & TIME — Batch 2
// ============================================================================

// DAYS360(start_date, end_date, [method])
QVariant FormulaEngine::funcDAYS360(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    QDate start = parseDate(args[0]);
    QDate end = parseDate(args[1]);
    if (!start.isValid() || !end.isValid()) return QVariant("#VALUE!");
    bool euMethod = args.size() > 2 ? toBoolean(args[2]) : false;

    int d1 = start.day(), m1 = start.month(), y1 = start.year();
    int d2 = end.day(), m2 = end.month(), y2 = end.year();

    if (euMethod) {
        if (d1 > 30) d1 = 30;
        if (d2 > 30) d2 = 30;
    } else {
        if (d1 == 31) d1 = 30;
        if (d2 == 31 && d1 >= 30) d2 = 30;
    }
    return (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1);
}

// NETWORKDAYS.INTL(start, end, [weekend], [holidays])
QVariant FormulaEngine::funcNETWORKDAYS_INTL(const std::vector<QVariant>& args) {
    // Simplified: same as NETWORKDAYS (weekend = Sat+Sun)
    return funcNETWORKDAYS(args);
}

// WORKDAY.INTL — simplified
QVariant FormulaEngine::funcWORKDAY_INTL(const std::vector<QVariant>& args) {
    return funcWORKDAY(args);
}

// ============================================================================
// INFORMATION — Batch 2
// ============================================================================

// ISFORMULA(reference)
QVariant FormulaEngine::funcISFORMULA(const std::vector<QVariant>& args) {
    if (args.empty()) return false;
    // Check if the referenced cell has a formula
    // In practice, this needs cell reference context; simplified:
    QString s = toString(args[0]);
    return s.startsWith("=");
}

// ISREF(value) — always false in formula context (simplified)
QVariant FormulaEngine::funcISREF(const std::vector<QVariant>& args) {
    if (args.empty()) return false;
    return args[0].canConvert<CellRange>();
}
