#include "DocumentTheme.h"
#include <QStringList>
#include <cmath>

// ===== Tint Algorithm (Excel-compatible, HSL space) =====
// tint > 0: interpolate luminance toward white: lum' = lum * (1 - tint) + tint
// tint < 0: interpolate luminance toward black: lum' = lum * (1 + tint)

QColor DocumentTheme::applyTint(const QColor& base, double tintFactor) {
    if (std::abs(tintFactor) < 0.001) return base;

    float h, s, l, a;
    base.getHslF(&h, &s, &l, &a);

    if (tintFactor > 0.0) {
        l = static_cast<float>(l * (1.0 - tintFactor) + tintFactor);
    } else {
        l = static_cast<float>(l * (1.0 + tintFactor));
    }
    l = std::clamp(l, 0.0f, 1.0f);

    QColor result;
    result.setHslF(h, s, l, a);
    return result;
}

// ===== Theme color string parsing =====

bool DocumentTheme::isThemeColor(const QString& colorStr) {
    return colorStr.startsWith(QLatin1String("theme:"));
}

QString DocumentTheme::makeThemeColorStr(int themeIndex, double tint) {
    return QString("theme:%1:%2").arg(themeIndex).arg(tint, 0, 'g', 4);
}

QColor DocumentTheme::resolveColor(const QString& colorStr) const {
    if (!isThemeColor(colorStr)) return QColor(); // invalid

    // Parse "theme:<index>:<tint>"
    QStringList parts = colorStr.mid(6).split(':');
    if (parts.size() < 2) return QColor();

    bool okIdx = false, okTint = false;
    int index = parts[0].toInt(&okIdx);
    double tint = parts[1].toDouble(&okTint);
    if (!okIdx || !okTint || index < 0 || index > 9) return QColor();

    return applyTint(colors[index], tint);
}

QColor DocumentTheme::resolveAnyColor(const QString& colorStr) const {
    if (isThemeColor(colorStr)) {
        return resolveColor(colorStr);
    }
    return QColor(colorStr);
}

// ===== Theme color display names =====

static const char* kColorSlotNames[] = {
    "Dark 1", "Light 1", "Dark 2", "Light 2",
    "Accent 1", "Accent 2", "Accent 3", "Accent 4", "Accent 5", "Accent 6"
};

static const char* kTintDescriptions[] = {
    "",                 // 0.0 = base
    ", 80% Lighter",    // 0.8
    ", 60% Lighter",    // 0.6
    ", 40% Lighter",    // 0.4
    ", 25% Darker",     // -0.25
    ", 50% Darker"      // -0.5
};

QString themeColorName(int index, double tint) {
    if (index < 0 || index > 9) return "Unknown";
    QString name = kColorSlotNames[index];

    // Find matching tint description
    for (int i = 0; i < kThemeTintCount; ++i) {
        if (std::abs(tint - kThemeTints[i]) < 0.01) {
            return name + kTintDescriptions[i];
        }
    }
    // Custom tint
    if (tint > 0.0) {
        return name + QString(", %1% Lighter").arg(static_cast<int>(tint * 100));
    } else if (tint < 0.0) {
        return name + QString(", %1% Darker").arg(static_cast<int>(-tint * 100));
    }
    return name;
}

// ===== 8 Predefined Document Themes =====

static DocumentTheme createOfficeTheme() {
    DocumentTheme t;
    t.id = "office";
    t.displayName = "Office";
    t.colors = {{
        QColor("#000000"), // Dark 1
        QColor("#FFFFFF"), // Light 1
        QColor("#44546A"), // Dark 2
        QColor("#E7E6E6"), // Light 2
        QColor("#4472C4"), // Accent 1
        QColor("#ED7D31"), // Accent 2
        QColor("#A5A5A5"), // Accent 3
        QColor("#FFC000"), // Accent 4
        QColor("#5B9BD5"), // Accent 5
        QColor("#70AD47"), // Accent 6
    }};
    return t;
}

static DocumentTheme createBlueWarmTheme() {
    DocumentTheme t;
    t.id = "blue-warm";
    t.displayName = "Blue Warm";
    t.colors = {{
        QColor("#000000"),
        QColor("#FFFFFF"),
        QColor("#242852"),
        QColor("#DFDEE3"),
        QColor("#4A66AC"),
        QColor("#629DD1"),
        QColor("#297FD5"),
        QColor("#7F8FA9"),
        QColor("#5AA2AE"),
        QColor("#9D90A0"),
    }};
    return t;
}

static DocumentTheme createCoolTheme() {
    DocumentTheme t;
    t.id = "cool";
    t.displayName = "Cool";
    t.colors = {{
        QColor("#000000"),
        QColor("#FFFFFF"),
        QColor("#1F497D"),
        QColor("#EEECE1"),
        QColor("#4F81BD"),
        QColor("#C0504D"),
        QColor("#9BBB59"),
        QColor("#8064A2"),
        QColor("#4BACC6"),
        QColor("#F79646"),
    }};
    return t;
}

static DocumentTheme createWarmTheme() {
    DocumentTheme t;
    t.id = "warm";
    t.displayName = "Warm";
    t.colors = {{
        QColor("#000000"),
        QColor("#FFFFFF"),
        QColor("#4E3B30"),
        QColor("#FFF2E2"),
        QColor("#C0504D"),
        QColor("#F79646"),
        QColor("#9BBB59"),
        QColor("#8064A2"),
        QColor("#4BACC6"),
        QColor("#4F81BD"),
    }};
    return t;
}

static DocumentTheme createGreenTheme() {
    DocumentTheme t;
    t.id = "green";
    t.displayName = "Green";
    t.colors = {{
        QColor("#000000"),
        QColor("#FFFFFF"),
        QColor("#2D572C"),
        QColor("#D5EDDB"),
        QColor("#549E39"),
        QColor("#8AB833"),
        QColor("#C0CF3A"),
        QColor("#029676"),
        QColor("#4AB5C4"),
        QColor("#0989B1"),
    }};
    return t;
}

static DocumentTheme createMarqueeTheme() {
    DocumentTheme t;
    t.id = "marquee";
    t.displayName = "Marquee";
    t.colors = {{
        QColor("#000000"),
        QColor("#FFFFFF"),
        QColor("#3B3059"),
        QColor("#E8E3F1"),
        QColor("#7B2D8D"),
        QColor("#F5C201"),
        QColor("#E73C53"),
        QColor("#1B9AA0"),
        QColor("#6B4E99"),
        QColor("#F09E00"),
    }};
    return t;
}

static DocumentTheme createSlipstreamTheme() {
    DocumentTheme t;
    t.id = "slipstream";
    t.displayName = "Slipstream";
    t.colors = {{
        QColor("#000000"),
        QColor("#FFFFFF"),
        QColor("#212745"),
        QColor("#B4DCFA"),
        QColor("#4E67C8"),
        QColor("#5ECCF3"),
        QColor("#A7EA52"),
        QColor("#5DCEAF"),
        QColor("#FF8021"),
        QColor("#F14124"),
    }};
    return t;
}

static DocumentTheme createGrayscaleTheme() {
    DocumentTheme t;
    t.id = "grayscale";
    t.displayName = "Grayscale";
    t.colors = {{
        QColor("#000000"),
        QColor("#FFFFFF"),
        QColor("#333333"),
        QColor("#E0E0E0"),
        QColor("#666666"),
        QColor("#999999"),
        QColor("#808080"),
        QColor("#B3B3B3"),
        QColor("#4D4D4D"),
        QColor("#CCCCCC"),
    }};
    return t;
}

// ===== Registry =====

std::vector<DocumentTheme> getBuiltinDocumentThemes() {
    return {
        createOfficeTheme(),
        createBlueWarmTheme(),
        createCoolTheme(),
        createWarmTheme(),
        createGreenTheme(),
        createMarqueeTheme(),
        createSlipstreamTheme(),
        createGrayscaleTheme(),
    };
}

const DocumentTheme& defaultDocumentTheme() {
    static DocumentTheme office = createOfficeTheme();
    return office;
}
