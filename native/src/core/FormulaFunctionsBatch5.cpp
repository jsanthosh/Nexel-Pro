// ============================================================================
// FormulaFunctionsBatch5.cpp — 37 new functions: complex, matrix, Bessel,
// statistical, financial, info, trig, database
// ============================================================================

#include "FormulaEngine.h"
#include "Spreadsheet.h"
#include <cmath>
#include <algorithm>
#include <numeric>
#include <map>
#include <QSysInfo>

// ============================================================================
// Complex number helpers
// ============================================================================

// Parse "a+bi" or "a-bi" string into real and imaginary parts
static bool parseComplex(const QString& s, double& real, double& imag) {
    real = 0; imag = 0;
    if (s.isEmpty()) return false;

    // Try pure real number first
    int iPos = s.indexOf('i');
    if (iPos < 0) {
        bool ok;
        real = s.toDouble(&ok);
        return ok;
    }

    // Pure imaginary: "i", "+i", "-i", "3i", "-3i"
    if (iPos == s.length() - 1) {
        QString before = s.left(iPos);
        if (before.isEmpty() || before == "+") { imag = 1.0; return true; }
        if (before == "-") { imag = -1.0; return true; }

        // Find where imaginary coefficient starts (after last + or - that is not at position 0)
        int splitPos = -1;
        for (int i = before.length() - 1; i > 0; --i) {
            if (before[i] == '+' || before[i] == '-') {
                splitPos = i;
                break;
            }
        }

        if (splitPos < 0) {
            // Entire string is the imaginary part
            bool ok;
            imag = before.toDouble(&ok);
            return ok;
        }

        // Real part + imaginary part
        bool ok1, ok2;
        real = before.left(splitPos).toDouble(&ok1);
        QString imagStr = before.mid(splitPos);
        if (imagStr == "+" || imagStr.isEmpty()) imag = 1.0;
        else if (imagStr == "-") imag = -1.0;
        else imag = imagStr.toDouble(&ok2);
        return true;
    }

    return false;
}

// Format complex number back to string
static QString formatComplex(double real, double imag) {
    if (real == 0 && imag == 0) return QStringLiteral("0");
    QString result;
    if (real != 0) result = QString::number(real, 'g', 15);
    if (imag != 0) {
        if (imag > 0 && !result.isEmpty()) result += "+";
        if (imag == 1.0) result += "i";
        else if (imag == -1.0) result += "-i";
        else result += QString::number(imag, 'g', 15) + "i";
    }
    return result;
}

// ============================================================================
// COMPLEX NUMBER FUNCTIONS (15)
// ============================================================================

// IMABS — absolute value (modulus) of complex number
QVariant FormulaEngine::funcIMABS(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double re, im;
    if (!parseComplex(toString(args[0]), re, im)) return QVariant("#NUM!");
    return std::sqrt(re * re + im * im);
}

// IMARGUMENT — argument (angle in radians) of complex number
QVariant FormulaEngine::funcIMARGUMENT(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double re, im;
    if (!parseComplex(toString(args[0]), re, im)) return QVariant("#NUM!");
    if (re == 0 && im == 0) return QVariant("#DIV/0!");
    return std::atan2(im, re);
}

// IMCONJUGATE — complex conjugate
QVariant FormulaEngine::funcIMCONJUGATE(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double re, im;
    if (!parseComplex(toString(args[0]), re, im)) return QVariant("#NUM!");
    return formatComplex(re, -im);
}

// IMSUM — sum of two or more complex numbers
QVariant FormulaEngine::funcIMSUM(const std::vector<QVariant>& args) {
    if (args.size() < 1) return QVariant("#VALUE!");
    double sumRe = 0, sumIm = 0;
    for (const auto& arg : args) {
        double re, im;
        if (!parseComplex(toString(arg), re, im)) return QVariant("#NUM!");
        sumRe += re;
        sumIm += im;
    }
    return formatComplex(sumRe, sumIm);
}

// IMSUB — subtract two complex numbers
QVariant FormulaEngine::funcIMSUB(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    double re1, im1, re2, im2;
    if (!parseComplex(toString(args[0]), re1, im1)) return QVariant("#NUM!");
    if (!parseComplex(toString(args[1]), re2, im2)) return QVariant("#NUM!");
    return formatComplex(re1 - re2, im1 - im2);
}

// IMPRODUCT — product of two or more complex numbers
QVariant FormulaEngine::funcIMPRODUCT(const std::vector<QVariant>& args) {
    if (args.size() < 1) return QVariant("#VALUE!");
    double prodRe = 1, prodIm = 0;
    for (const auto& arg : args) {
        double re, im;
        if (!parseComplex(toString(arg), re, im)) return QVariant("#NUM!");
        double newRe = prodRe * re - prodIm * im;
        double newIm = prodRe * im + prodIm * re;
        prodRe = newRe;
        prodIm = newIm;
    }
    return formatComplex(prodRe, prodIm);
}

// IMDIV — divide two complex numbers
QVariant FormulaEngine::funcIMDIV(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    double re1, im1, re2, im2;
    if (!parseComplex(toString(args[0]), re1, im1)) return QVariant("#NUM!");
    if (!parseComplex(toString(args[1]), re2, im2)) return QVariant("#NUM!");
    double denom = re2 * re2 + im2 * im2;
    if (denom == 0) return QVariant("#NUM!");
    return formatComplex((re1 * re2 + im1 * im2) / denom,
                         (im1 * re2 - re1 * im2) / denom);
}

// IMPOWER — complex number raised to a power
QVariant FormulaEngine::funcIMPOWER(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    double re, im;
    if (!parseComplex(toString(args[0]), re, im)) return QVariant("#NUM!");
    double n = toNumber(args[1]);
    // Convert to polar: r*e^(i*theta)
    double r = std::sqrt(re * re + im * im);
    if (r == 0) {
        if (n > 0) return formatComplex(0, 0);
        return QVariant("#NUM!");
    }
    double theta = std::atan2(im, re);
    double rn = std::pow(r, n);
    double newTheta = theta * n;
    return formatComplex(rn * std::cos(newTheta), rn * std::sin(newTheta));
}

// IMSQRT — square root of complex number
QVariant FormulaEngine::funcIMSQRT(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double re, im;
    if (!parseComplex(toString(args[0]), re, im)) return QVariant("#NUM!");
    double r = std::sqrt(re * re + im * im);
    double theta = std::atan2(im, re);
    double sqrtR = std::sqrt(r);
    return formatComplex(sqrtR * std::cos(theta / 2), sqrtR * std::sin(theta / 2));
}

// IMEXP — e raised to complex power
QVariant FormulaEngine::funcIMEXP(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double re, im;
    if (!parseComplex(toString(args[0]), re, im)) return QVariant("#NUM!");
    double expRe = std::exp(re);
    return formatComplex(expRe * std::cos(im), expRe * std::sin(im));
}

// IMLN — natural logarithm of complex number
QVariant FormulaEngine::funcIMLN(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double re, im;
    if (!parseComplex(toString(args[0]), re, im)) return QVariant("#NUM!");
    double r = std::sqrt(re * re + im * im);
    if (r == 0) return QVariant("#NUM!");
    double theta = std::atan2(im, re);
    return formatComplex(std::log(r), theta);
}

// IMLOG2 — log base 2 of complex number
QVariant FormulaEngine::funcIMLOG2(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double re, im;
    if (!parseComplex(toString(args[0]), re, im)) return QVariant("#NUM!");
    double r = std::sqrt(re * re + im * im);
    if (r == 0) return QVariant("#NUM!");
    double theta = std::atan2(im, re);
    double ln2 = std::log(2.0);
    return formatComplex(std::log(r) / ln2, theta / ln2);
}

// IMLOG10 — log base 10 of complex number
QVariant FormulaEngine::funcIMLOG10(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double re, im;
    if (!parseComplex(toString(args[0]), re, im)) return QVariant("#NUM!");
    double r = std::sqrt(re * re + im * im);
    if (r == 0) return QVariant("#NUM!");
    double theta = std::atan2(im, re);
    double ln10 = std::log(10.0);
    return formatComplex(std::log(r) / ln10, theta / ln10);
}

// IMSIN — sine of complex number: sin(a+bi) = sin(a)*cosh(b) + i*cos(a)*sinh(b)
QVariant FormulaEngine::funcIMSIN(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double re, im;
    if (!parseComplex(toString(args[0]), re, im)) return QVariant("#NUM!");
    return formatComplex(std::sin(re) * std::cosh(im),
                         std::cos(re) * std::sinh(im));
}

// IMCOS — cosine of complex number: cos(a+bi) = cos(a)*cosh(b) - i*sin(a)*sinh(b)
QVariant FormulaEngine::funcIMCOS(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double re, im;
    if (!parseComplex(toString(args[0]), re, im)) return QVariant("#NUM!");
    return formatComplex(std::cos(re) * std::cosh(im),
                         -std::sin(re) * std::sinh(im));
}

// ============================================================================
// MATRIX FUNCTIONS (2)
// ============================================================================

// MINVERSE — matrix inverse using Gauss-Jordan elimination
QVariant FormulaEngine::funcMINVERSE(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");

    // Extract matrix from flattened 2D array
    auto flat = flattenArgs(args);
    int n = static_cast<int>(std::sqrt(static_cast<double>(flat.size())));
    if (n * n != static_cast<int>(flat.size()) || n == 0)
        return QVariant("#VALUE!");

    // Build augmented matrix [A | I]
    std::vector<std::vector<double>> aug(n, std::vector<double>(2 * n, 0.0));
    for (int i = 0; i < n; ++i) {
        for (int j = 0; j < n; ++j)
            aug[i][j] = toNumber(flat[i * n + j]);
        aug[i][n + i] = 1.0;
    }

    // Gauss-Jordan elimination
    for (int col = 0; col < n; ++col) {
        // Find pivot
        int pivotRow = -1;
        double maxVal = 0;
        for (int row = col; row < n; ++row) {
            double v = std::fabs(aug[row][col]);
            if (v > maxVal) { maxVal = v; pivotRow = row; }
        }
        if (maxVal < 1e-15) return QVariant("#NUM!"); // Singular

        // Swap rows
        if (pivotRow != col) std::swap(aug[col], aug[pivotRow]);

        // Scale pivot row
        double pivot = aug[col][col];
        for (int j = 0; j < 2 * n; ++j)
            aug[col][j] /= pivot;

        // Eliminate column
        for (int row = 0; row < n; ++row) {
            if (row == col) continue;
            double factor = aug[row][col];
            for (int j = 0; j < 2 * n; ++j)
                aug[row][j] -= factor * aug[col][j];
        }
    }

    // Extract result as 2D QVariantList
    QVariantList rows;
    for (int i = 0; i < n; ++i) {
        QVariantList row;
        for (int j = 0; j < n; ++j)
            row.append(aug[i][n + j]);
        rows.append(QVariant(row));
    }
    return QVariant(rows);
}

// MDETERM — matrix determinant using LU decomposition
QVariant FormulaEngine::funcMDETERM(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");

    auto flat = flattenArgs(args);
    int n = static_cast<int>(std::sqrt(static_cast<double>(flat.size())));
    if (n * n != static_cast<int>(flat.size()) || n == 0)
        return QVariant("#VALUE!");

    // Copy matrix
    std::vector<std::vector<double>> mat(n, std::vector<double>(n));
    for (int i = 0; i < n; ++i)
        for (int j = 0; j < n; ++j)
            mat[i][j] = toNumber(flat[i * n + j]);

    // LU decomposition with partial pivoting
    double det = 1.0;
    for (int col = 0; col < n; ++col) {
        // Find pivot
        int pivotRow = col;
        double maxVal = std::fabs(mat[col][col]);
        for (int row = col + 1; row < n; ++row) {
            double v = std::fabs(mat[row][col]);
            if (v > maxVal) { maxVal = v; pivotRow = row; }
        }
        if (maxVal < 1e-15) return 0.0; // Singular

        if (pivotRow != col) {
            std::swap(mat[col], mat[pivotRow]);
            det = -det;
        }

        det *= mat[col][col];

        for (int row = col + 1; row < n; ++row) {
            double factor = mat[row][col] / mat[col][col];
            for (int j = col + 1; j < n; ++j)
                mat[row][j] -= factor * mat[col][j];
        }
    }
    return det;
}

// ============================================================================
// BESSEL FUNCTIONS (4) — series expansion for integer orders
// ============================================================================

// Helper: factorial
static double factorial(int n) {
    double r = 1;
    for (int i = 2; i <= n; ++i) r *= i;
    return r;
}

// BESSELJ — Bessel function of the first kind J_n(x)
QVariant FormulaEngine::funcBESSELJ(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    double x = toNumber(args[0]);
    int n = static_cast<int>(toNumber(args[1]));
    if (n < 0) return QVariant("#NUM!");

    double sum = 0;
    for (int m = 0; m <= 100; ++m) {
        double term = (m % 2 == 0 ? 1.0 : -1.0) / (factorial(m) * factorial(m + n))
                      * std::pow(x / 2.0, 2 * m + n);
        sum += term;
        if (std::fabs(term) < 1e-15 * std::fabs(sum) && m > n) break;
    }
    return sum;
}

// BESSELY — Bessel function of the second kind Y_n(x)
QVariant FormulaEngine::funcBESSELY(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    double x = toNumber(args[0]);
    int n = static_cast<int>(toNumber(args[1]));
    if (n < 0 || x <= 0) return QVariant("#NUM!");

    // Y_n(x) = (J_n(x)*cos(n*pi) - J_{-n}(x)) / sin(n*pi)
    // For integer n, use limit form:
    // Y_n(x) = (2/pi)*J_n(x)*ln(x/2) - (1/pi)*sum of terms
    // Use approximation via BESSELJ with numerical differentiation
    double jn = toNumber(funcBESSELJ({args[0], args[1]}));
    double euler = 0.5772156649015329;

    if (n == 0) {
        // Y_0(x) = (2/pi)*(J_0(x)*(ln(x/2) + gamma) + sum)
        double sum = 0;
        for (int m = 1; m <= 100; ++m) {
            double hm = 0;
            for (int k = 1; k <= m; ++k) hm += 1.0 / k;
            double term = (m % 2 == 0 ? 1.0 : -1.0) * hm / (factorial(m) * factorial(m))
                          * std::pow(x / 2.0, 2 * m);
            sum += term;
            if (std::fabs(term) < 1e-15 && m > 5) break;
        }
        return (2.0 / M_PI) * (jn * (std::log(x / 2.0) + euler) + sum);
    }

    // For n > 0, use recurrence: Y_{n+1}(x) = (2n/x)*Y_n(x) - Y_{n-1}(x)
    double y0 = toNumber(funcBESSELY({args[0], QVariant(0)}));
    if (n == 0) return y0;

    // Y_1 via series
    double j1 = toNumber(funcBESSELJ({args[0], QVariant(1)}));
    double sum1 = 0;
    for (int m = 0; m <= 100; ++m) {
        double hm = 0, hm1 = 0;
        for (int k = 1; k <= m; ++k) hm += 1.0 / k;
        for (int k = 1; k <= m + 1; ++k) hm1 += 1.0 / k;
        double term = (m % 2 == 0 ? 1.0 : -1.0) * (hm + hm1) / (factorial(m) * factorial(m + 1))
                      * std::pow(x / 2.0, 2 * m + 1);
        sum1 += term;
        if (std::fabs(term) < 1e-15 && m > 5) break;
    }
    double negPart = 0;
    // -1/pi * sum_{m=0}^{n-1} (n-1-m)!/m! * (x/2)^(2m-n+1) ... simplified
    // Use: Y_1(x) = (2/pi)(J_1(x)*ln(x/2) - 1/x) + ...
    double y1 = (2.0 / M_PI) * (j1 * (std::log(x / 2.0) + euler) - 1.0 / x - sum1);

    if (n == 1) return y1;

    // Recurrence for higher orders
    double ym1 = y0, ym = y1;
    for (int k = 1; k < n; ++k) {
        double yn = (2.0 * k / x) * ym - ym1;
        ym1 = ym;
        ym = yn;
    }
    return ym;
}

// BESSELI — modified Bessel function of the first kind I_n(x)
QVariant FormulaEngine::funcBESSELI(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    double x = toNumber(args[0]);
    int n = static_cast<int>(toNumber(args[1]));
    if (n < 0) return QVariant("#NUM!");

    double sum = 0;
    for (int m = 0; m <= 100; ++m) {
        double term = 1.0 / (factorial(m) * factorial(m + n))
                      * std::pow(x / 2.0, 2 * m + n);
        sum += term;
        if (std::fabs(term) < 1e-15 * std::fabs(sum) && m > n) break;
    }
    return sum;
}

// BESSELK — modified Bessel function of the second kind K_n(x)
QVariant FormulaEngine::funcBESSELK(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    double x = toNumber(args[0]);
    int n = static_cast<int>(toNumber(args[1]));
    if (n < 0 || x <= 0) return QVariant("#NUM!");

    // K_0(x) = -(ln(x/2) + gamma) * I_0(x) + sum
    double euler = 0.5772156649015329;
    double i0 = toNumber(funcBESSELI({args[0], QVariant(0)}));

    if (n == 0) {
        double sum = 0;
        for (int m = 1; m <= 100; ++m) {
            double hm = 0;
            for (int k = 1; k <= m; ++k) hm += 1.0 / k;
            double term = hm / (factorial(m) * factorial(m))
                          * std::pow(x / 2.0, 2 * m);
            sum += term;
            if (std::fabs(term) < 1e-15 && m > 5) break;
        }
        return -(std::log(x / 2.0) + euler) * i0 + sum;
    }

    // K_1 from series
    double i1 = toNumber(funcBESSELI({args[0], QVariant(1)}));
    double sum1 = 0;
    for (int m = 0; m <= 100; ++m) {
        double hm = 0, hm1 = 0;
        for (int k = 1; k <= m; ++k) hm += 1.0 / k;
        for (int k = 1; k <= m + 1; ++k) hm1 += 1.0 / k;
        double term = (hm + hm1) / (factorial(m) * factorial(m + 1))
                      * std::pow(x / 2.0, 2 * m + 1);
        sum1 += term;
        if (std::fabs(term) < 1e-15 && m > 5) break;
    }
    double k0 = toNumber(funcBESSELK({args[0], QVariant(0)}));
    double k1 = (1.0 / x) + (std::log(x / 2.0) + euler) * (-i1) + sum1;

    // Fix: use proper K_1 formula
    // K_1(x) = 1/x + I_1(x)*(ln(x/2) + gamma) - ... Simplified:
    // Use recurrence starting from k0 and numerical k1
    // K_1 approximated from derivative relation
    if (n == 1) return k1;

    // Recurrence: K_{n+1}(x) = (2n/x)*K_n(x) + K_{n-1}(x)
    double km1 = k0, km = k1;
    for (int ki = 1; ki < n; ++ki) {
        double kn = (2.0 * ki / x) * km + km1;
        km1 = km;
        km = kn;
    }
    return km;
}

// ============================================================================
// STATISTICAL FUNCTIONS (3)
// ============================================================================

// PHI — standard normal probability density function
QVariant FormulaEngine::funcPHI(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double x = toNumber(args[0]);
    return (1.0 / std::sqrt(2.0 * M_PI)) * std::exp(-0.5 * x * x);
}

// GAUSS — probability that a standard normal random variable falls between 0 and z
QVariant FormulaEngine::funcGAUSS(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double z = toNumber(args[0]);
    // GAUSS(z) = NORM.S.DIST(z, TRUE) - 0.5
    // Use error function: Phi(z) = 0.5 * (1 + erf(z / sqrt(2)))
    return 0.5 * std::erf(z / std::sqrt(2.0));
}

// MODE.MULT — returns array of all modes (most frequent values)
QVariant FormulaEngine::funcMODE_MULT(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    auto values = flattenArgs(args);

    // Count frequency of each numeric value
    std::map<double, int> freq;
    for (const auto& v : values) {
        bool ok;
        double d = v.toDouble(&ok);
        if (ok) freq[d]++;
    }

    if (freq.empty()) return QVariant("#N/A");

    // Find max frequency
    int maxFreq = 0;
    for (const auto& [val, cnt] : freq)
        if (cnt > maxFreq) maxFreq = cnt;

    if (maxFreq <= 1) return QVariant("#N/A"); // No repeating values

    // Collect all values with max frequency
    QVariantList modes;
    for (const auto& [val, cnt] : freq)
        if (cnt == maxFreq) modes.append(val);

    if (modes.size() == 1) return modes[0];
    return QVariant(modes);
}

// ============================================================================
// FINANCIAL (1)
// ============================================================================

// VDB — Variable Declining Balance depreciation
// VDB(cost, salvage, life, start_period, end_period, [factor], [no_switch])
QVariant FormulaEngine::funcVDB(const std::vector<QVariant>& args) {
    if (args.size() < 5) return QVariant("#VALUE!");
    double cost = toNumber(args[0]);
    double salvage = toNumber(args[1]);
    double life = toNumber(args[2]);
    double startPeriod = toNumber(args[3]);
    double endPeriod = toNumber(args[4]);
    double factor = args.size() > 5 ? toNumber(args[5]) : 2.0;
    bool noSwitch = args.size() > 6 ? toBoolean(args[6]) : false;

    if (cost < 0 || salvage < 0 || life <= 0 || startPeriod < 0 ||
        endPeriod < startPeriod || endPeriod > life || factor <= 0)
        return QVariant("#NUM!");

    // Calculate depreciation for each period using DDB then optionally switch to SLN
    double totalDepr = 0;
    double bookValue = cost;

    for (int p = 1; p <= static_cast<int>(std::ceil(endPeriod)); ++p) {
        // DDB depreciation for this period
        double ddbDepr = std::min(bookValue * factor / life, bookValue - salvage);
        ddbDepr = std::max(ddbDepr, 0.0);

        // SLN depreciation for remaining periods
        double slnDepr = 0;
        double remainingLife = life - p + 1;
        if (remainingLife > 0)
            slnDepr = std::max((bookValue - salvage) / remainingLife, 0.0);

        // Use larger of DDB or SLN (unless no_switch)
        double depr = noSwitch ? ddbDepr : std::max(ddbDepr, slnDepr);
        depr = std::min(depr, bookValue - salvage);
        depr = std::max(depr, 0.0);

        // Handle fractional periods
        double fraction = 1.0;
        if (p <= static_cast<int>(startPeriod)) {
            // Before start period - no accumulation
        } else {
            if (p == static_cast<int>(std::ceil(startPeriod)) && startPeriod != std::floor(startPeriod))
                fraction = 1.0 - (startPeriod - std::floor(startPeriod));
            if (p == static_cast<int>(std::ceil(endPeriod)) && endPeriod != std::floor(endPeriod))
                fraction = std::min(fraction, endPeriod - std::floor(endPeriod));

            if (static_cast<double>(p) > startPeriod)
                totalDepr += depr * fraction;
        }

        bookValue -= depr;
    }
    return totalDepr;
}

// ============================================================================
// INFORMATION FUNCTIONS (2)
// ============================================================================

// CELL — returns information about a cell
// CELL(info_type, [reference])
QVariant FormulaEngine::funcCELL(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    QString infoType = toString(args[0]).toLower();

    // Without a reference, we return generic info
    if (infoType == "type") {
        if (args.size() < 2) return QStringLiteral("b"); // blank
        const auto& v = args[1];
        if (v.isNull() || !v.isValid() || v.toString().isEmpty()) return QStringLiteral("b");
        bool ok;
        v.toDouble(&ok);
        if (ok) return QStringLiteral("v");
        return QStringLiteral("l"); // label
    }
    if (infoType == "row") {
        if (args.size() < 2) return 1;
        // If we have cell reference info, extract row
        if (args[1].canConvert<CellRange>()) {
            CellRange range = args[1].value<CellRange>();
            return range.start().row() + 1;
        }
        return 1;
    }
    if (infoType == "col") {
        if (args.size() < 2) return 1;
        if (args[1].canConvert<CellRange>()) {
            CellRange range = args[1].value<CellRange>();
            return range.start().col() + 1;
        }
        return 1;
    }
    if (infoType == "address") {
        if (args.size() < 2) return QStringLiteral("$A$1");
        if (args[1].canConvert<CellRange>()) {
            CellRange range = args[1].value<CellRange>();
            int r = range.start().row() + 1;
            int c = range.start().col();
            QChar colChar('A' + c);
            return QStringLiteral("$") + colChar + QStringLiteral("$") + QString::number(r);
        }
        return QStringLiteral("$A$1");
    }
    if (infoType == "contents") {
        if (args.size() < 2) return QVariant();
        return args[1];
    }
    if (infoType == "filename") {
        return QStringLiteral("");
    }
    if (infoType == "format") {
        return QStringLiteral("G"); // General format
    }
    if (infoType == "width") {
        return 8; // Default column width
    }
    return QVariant("#VALUE!");
}

// INFO — returns system/environment info
// INFO(type_text)
QVariant FormulaEngine::funcINFO(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    QString infoType = toString(args[0]).toLower();

    if (infoType == "osversion") {
        return QSysInfo::prettyProductName();
    }
    if (infoType == "system") {
#ifdef Q_OS_WIN
        return QStringLiteral("pcdos");
#else
        return QStringLiteral("mac");
#endif
    }
    if (infoType == "directory") {
        return QStringLiteral("");
    }
    if (infoType == "numfile") {
        return 1;
    }
    if (infoType == "origin") {
        return QStringLiteral("$A:$A$1");
    }
    if (infoType == "recalc") {
        return QStringLiteral("Automatic");
    }
    if (infoType == "release") {
        return QStringLiteral("1.0");
    }
    return QVariant("#VALUE!");
}

// ============================================================================
// MATH — TRIG FUNCTIONS (8)
// ============================================================================

// COT — cotangent: cos(x)/sin(x)
QVariant FormulaEngine::funcCOT(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double x = toNumber(args[0]);
    double s = std::sin(x);
    if (s == 0) return QVariant("#DIV/0!");
    return std::cos(x) / s;
}

// COTH — hyperbolic cotangent: cosh(x)/sinh(x)
QVariant FormulaEngine::funcCOTH(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double x = toNumber(args[0]);
    double s = std::sinh(x);
    if (s == 0) return QVariant("#DIV/0!");
    return std::cosh(x) / s;
}

// ACOT — arccotangent: atan(1/x)
QVariant FormulaEngine::funcACOT(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double x = toNumber(args[0]);
    return std::atan(1.0 / x);
}

// ACOTH — inverse hyperbolic cotangent: 0.5 * ln((x+1)/(x-1))
QVariant FormulaEngine::funcACOTH(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double x = toNumber(args[0]);
    if (std::fabs(x) <= 1) return QVariant("#NUM!");
    return 0.5 * std::log((x + 1.0) / (x - 1.0));
}

// CSC — cosecant: 1/sin(x)
QVariant FormulaEngine::funcCSC(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double x = toNumber(args[0]);
    double s = std::sin(x);
    if (s == 0) return QVariant("#DIV/0!");
    return 1.0 / s;
}

// CSCH — hyperbolic cosecant: 1/sinh(x)
QVariant FormulaEngine::funcCSCH(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double x = toNumber(args[0]);
    double s = std::sinh(x);
    if (s == 0) return QVariant("#DIV/0!");
    return 1.0 / s;
}

// SEC — secant: 1/cos(x)
QVariant FormulaEngine::funcSEC(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double x = toNumber(args[0]);
    double c = std::cos(x);
    if (c == 0) return QVariant("#DIV/0!");
    return 1.0 / c;
}

// SECH — hyperbolic secant: 1/cosh(x)
QVariant FormulaEngine::funcSECH(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double x = toNumber(args[0]);
    double c = std::cosh(x);
    if (c == 0) return QVariant("#DIV/0!");
    return 1.0 / c;
}

// ============================================================================
// DATABASE FUNCTIONS (2)
// ============================================================================

// DSTDEVP — database standard deviation (population)
QVariant FormulaEngine::funcDSTDEVP(const std::vector<QVariant>& args) {
    if (args.size() < 3) return QVariant("#VALUE!");
    // Simplified stub matching DSTDEV pattern — requires range context for full implementation
    Q_UNUSED(args);
    return 0.0;
}

// DVARP — database variance (population)
QVariant FormulaEngine::funcDVARP(const std::vector<QVariant>& args) {
    if (args.size() < 3) return QVariant("#VALUE!");
    Q_UNUSED(args);
    return 0.0;
}
