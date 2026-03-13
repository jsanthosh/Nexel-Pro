#ifndef THEME_H
#define THEME_H

#include <QColor>
#include <QString>
#include <QVector>
#include <QObject>

class QMainWindow;

struct NexelTheme {
    QString id;
    QString displayName;

    // Window & Chrome
    QColor windowBackground;
    QColor menuBarBackground;
    QColor menuBarText;
    QColor menuBarHover;
    QColor menuBackground;
    QColor menuBorder;
    QColor menuItemHover;
    QColor menuSeparator;
    QColor statusBarBackground;
    QColor statusBarText;

    // Toolbar
    QColor toolbarAccentStripe;
    QColor toolbarBackground;
    QColor toolbarBorder;
    QColor toolbarButtonText;
    QColor toolbarButtonHover;
    QColor toolbarButtonPressed;
    QColor toolbarButtonChecked;
    QColor toolbarButtonCheckedBorder;
    QColor toolbarInputBorder;
    QColor toolbarInputFocus;
    QColor toolbarSeparator;

    // Spreadsheet Grid
    QColor gridBackground;
    QColor gridLineColor;
    QColor headerBackground;
    QColor headerBorder;
    QColor headerText;
    QColor selectionTint;      // RGBA with alpha
    QColor focusBorderColor;
    QColor editorBorderColor;

    // Tab Bar (Bottom)
    QColor bottomBarBackground;
    QColor bottomBarBorder;
    QColor tabBackground;
    QColor tabBorder;
    QColor tabActiveBackground;
    QColor tabActiveIndicator;
    QColor tabHover;
    QColor addSheetButtonText;
    QColor addSheetButtonHover;

    // Formula Bar
    QColor formulaBarBackground;
    QColor formulaBarBorder;
    QColor formulaInputBorder;

    // Popups & Tooltips
    QColor popupBackground;
    QColor popupBorder;
    QColor popupItemSelected;
    QColor paramHintBackground;
    QColor paramHintBorder;

    // Chat Panel
    QColor chatHeaderGradientStart;
    QColor chatHeaderGradientEnd;
    QColor chatBackground;
    QColor chatUserBubble;
    QColor chatInputBackground;
    QColor chatInputBorder;
    QColor chatSendButton;
    QColor chatSendButtonHover;

    // Dock Widget
    QColor dockTitleBackground;
    QColor dockTitleText;

    // Dialogs
    QColor dialogBackground;
    QColor dialogInputBorder;
    QColor dialogButtonPrimary;
    QColor dialogButtonPrimaryHover;
    QColor dialogButtonPrimaryText;
    QColor dialogGroupBoxBorder;

    // Accent Colors
    QColor accentPrimary;
    QColor accentDark;
    QColor accentDarker;
    QColor accentLight;

    // Checkbox
    QColor checkboxChecked;
    QColor checkboxUncheckedBorder;

    // Freeze Line
    QColor freezeLineColor;

    // Text Colors
    QColor textPrimary;
    QColor textSecondary;
    QColor textMuted;

    // Selection Handle (chart/shape widgets)
    QColor selectionHandleColor;
};

class ThemeManager : public QObject {
    Q_OBJECT

public:
    static ThemeManager& instance();

    QVector<NexelTheme> availableThemes() const;
    const NexelTheme& currentTheme() const;
    bool isDarkTheme() const;
    void setTheme(const QString& themeId);
    void loadSavedTheme();
    void applyTheme(QMainWindow* window);
    void detectAndApplySystemTheme();

    // Shared stylesheet builders
    static QString dialogStylesheet();
    QString buildTabBarStylesheet() const;
    QString buildBottomBarStylesheet() const;
    QString buildAddSheetBtnStylesheet() const;

signals:
    void themeChanged(const NexelTheme& theme);

private:
    ThemeManager();
    void registerThemes();
    QString buildGlobalStylesheet() const;

    QVector<NexelTheme> m_themes;
    int m_currentIndex = 0;
};

#endif // THEME_H
