#include "Theme.h"
#include <QSettings>
#include <QMainWindow>
#include <QApplication>

// ── Theme Factory Functions ──

static NexelTheme createNexelGreen() {
    NexelTheme t;
    t.id = "nexel_green";
    t.displayName = "Nexel Green";

    // Window & Chrome
    t.windowBackground      = QColor("#F0F2F5");
    t.menuBarBackground      = QColor("#1B5E3B");
    t.menuBarText            = QColor("#FFFFFF");
    t.menuBarHover           = QColor("#155030");
    t.menuBackground         = QColor("#FFFFFF");
    t.menuBorder             = QColor("#D0D5DD");
    t.menuItemHover          = QColor("#E8F0FE");
    t.menuSeparator          = QColor("#E0E3E8");
    t.statusBarBackground    = QColor("#217346");
    t.statusBarText          = QColor("#FFFFFF");

    // Toolbar
    t.toolbarAccentStripe       = QColor("#1B5E3B");
    t.toolbarBackground         = QColor("#F0F2F5");
    t.toolbarBorder             = QColor("#D0D5DD");
    t.toolbarButtonText         = QColor("#344054");
    t.toolbarButtonHover        = QColor("#E8ECF0");
    t.toolbarButtonPressed      = QColor("#E0E3E8");
    t.toolbarButtonChecked      = QColor("#D6E4F0");
    t.toolbarButtonCheckedBorder = QColor("#4A90D9");
    t.toolbarInputBorder        = QColor("#D0D5DD");
    t.toolbarInputFocus         = QColor("#4A90D9");
    t.toolbarSeparator          = QColor("#E0E3E8");

    // Grid
    t.gridBackground     = QColor("#FFFFFF");
    t.gridLineColor      = QColor(218, 220, 224);
    t.headerBackground   = QColor("#F3F3F3");
    t.headerBorder       = QColor("#DADCE0");
    t.headerText         = QColor("#333333");
    t.selectionTint      = QColor(198, 217, 240, 60);
    t.focusBorderColor   = QColor(16, 124, 16);
    t.editorBorderColor  = QColor("#107C10");

    // Tab Bar
    t.bottomBarBackground   = QColor("#F3F3F3");
    t.bottomBarBorder       = QColor("#D0D0D0");
    t.tabBackground         = QColor("#E8E8E8");
    t.tabBorder             = QColor("#C8C8C8");
    t.tabActiveBackground   = QColor("#FFFFFF");
    t.tabActiveIndicator    = QColor("#217346");
    t.tabHover              = QColor("#D8D8D8");
    t.addSheetButtonText    = QColor("#555555");
    t.addSheetButtonHover   = QColor("#E0E0E0");

    // Formula Bar
    t.formulaBarBackground  = QColor("#FFFFFF");
    t.formulaBarBorder      = QColor("#E0E0E0");
    t.formulaInputBorder    = QColor("#D0D0D0");

    // Popups
    t.popupBackground       = QColor("#FFFFFF");
    t.popupBorder           = QColor("#C0C0C0");
    t.popupItemSelected     = QColor("#E8F0FE");
    t.paramHintBackground   = QColor("#FFF8DC");
    t.paramHintBorder       = QColor("#E0D8B0");

    // Chat Panel
    t.chatHeaderGradientStart = QColor("#1B5E3B");
    t.chatHeaderGradientEnd   = QColor("#2E8B57");
    t.chatBackground          = QColor("#ECE5DD");
    t.chatUserBubble          = QColor("#DCF8C6");
    t.chatInputBackground     = QColor("#F0F0F0");
    t.chatInputBorder         = QColor("#D9D9D9");
    t.chatSendButton          = QColor("#34A853");
    t.chatSendButtonHover     = QColor("#2D8C4E");

    // Dock Widget
    t.dockTitleBackground   = QColor("#1B5E3B");
    t.dockTitleText         = QColor("#FFFFFF");

    // Dialogs
    t.dialogBackground        = QColor("#F8FAFB");
    t.dialogInputBorder       = QColor("#D0D5DD");
    t.dialogButtonPrimary     = QColor("#217346");
    t.dialogButtonPrimaryHover = QColor("#1A5C38");
    t.dialogButtonPrimaryText = QColor("#FFFFFF");
    t.dialogGroupBoxBorder    = QColor("#D0D5DD");

    // Accents
    t.accentPrimary  = QColor("#34A853");
    t.accentDark     = QColor("#217346");
    t.accentDarker   = QColor("#1B5E3B");
    t.accentLight    = QColor("#E8F5E9");

    // Checkbox
    t.checkboxChecked        = QColor("#34A853");
    t.checkboxUncheckedBorder = QColor("#C4C7CC");

    // Freeze
    t.freezeLineColor = QColor("#808080");

    // Text
    t.textPrimary   = QColor("#333333");
    t.textSecondary = QColor("#555555");
    t.textMuted     = QColor("#94A3B8");

    // Selection Handle
    t.selectionHandleColor = QColor("#4A90D9");

    return t;
}

static NexelTheme createOceanBreeze() {
    NexelTheme t = createNexelGreen(); // start from defaults
    t.id = "ocean_breeze";
    t.displayName = "Ocean Breeze";

    t.menuBarBackground       = QColor("#2C5F7C");
    t.menuBarHover            = QColor("#234D65");
    t.statusBarBackground     = QColor("#3A7CA5");
    t.toolbarAccentStripe     = QColor("#2C5F7C");
    t.tabActiveIndicator      = QColor("#3A7CA5");
    t.focusBorderColor        = QColor("#3A7CA5");
    t.editorBorderColor       = QColor("#3A7CA5");
    t.windowBackground        = QColor("#EEF2F5");

    t.chatHeaderGradientStart = QColor("#2C5F7C");
    t.chatHeaderGradientEnd   = QColor("#4A90B0");
    t.chatBackground          = QColor("#E8F0F5");
    t.chatUserBubble          = QColor("#D4EAF7");
    t.chatSendButton          = QColor("#5BA4CF");
    t.chatSendButtonHover     = QColor("#4A90B0");

    t.dockTitleBackground     = QColor("#2C5F7C");

    t.dialogButtonPrimary     = QColor("#3A7CA5");
    t.dialogButtonPrimaryHover = QColor("#2C5F7C");

    t.accentPrimary  = QColor("#5BA4CF");
    t.accentDark     = QColor("#3A7CA5");
    t.accentDarker   = QColor("#2C5F7C");
    t.accentLight    = QColor("#D4EAF7");

    t.checkboxChecked         = QColor("#5BA4CF");
    t.selectionHandleColor    = QColor("#3A7CA5");

    t.menuItemHover           = QColor("#E0EDF5");
    t.toolbarButtonChecked    = QColor("#D4E6F1");
    t.popupItemSelected       = QColor("#E0EDF5");

    return t;
}

static NexelTheme createRoseQuartz() {
    NexelTheme t = createNexelGreen();
    t.id = "rose_quartz";
    t.displayName = "Rose Quartz";

    t.menuBarBackground       = QColor("#8A3F5C");
    t.menuBarHover            = QColor("#72334C");
    t.statusBarBackground     = QColor("#B05A78");
    t.toolbarAccentStripe     = QColor("#8A3F5C");
    t.tabActiveIndicator      = QColor("#B05A78");
    t.focusBorderColor        = QColor("#B05A78");
    t.editorBorderColor       = QColor("#B05A78");
    t.windowBackground        = QColor("#F5F0F2");

    t.chatHeaderGradientStart = QColor("#8A3F5C");
    t.chatHeaderGradientEnd   = QColor("#C06E8A");
    t.chatBackground          = QColor("#F5EBF0");
    t.chatUserBubble          = QColor("#F5E0E8");
    t.chatSendButton          = QColor("#D4829C");
    t.chatSendButtonHover     = QColor("#C06E8A");

    t.dockTitleBackground     = QColor("#8A3F5C");

    t.dialogButtonPrimary     = QColor("#B05A78");
    t.dialogButtonPrimaryHover = QColor("#8A3F5C");

    t.accentPrimary  = QColor("#D4829C");
    t.accentDark     = QColor("#B05A78");
    t.accentDarker   = QColor("#8A3F5C");
    t.accentLight    = QColor("#F5E0E8");

    t.checkboxChecked         = QColor("#D4829C");
    t.selectionHandleColor    = QColor("#B05A78");

    t.menuItemHover           = QColor("#F5E8EE");
    t.toolbarButtonChecked    = QColor("#F0D6E0");
    t.popupItemSelected       = QColor("#F5E8EE");

    return t;
}

static NexelTheme createLavenderMist() {
    NexelTheme t = createNexelGreen();
    t.id = "lavender_mist";
    t.displayName = "Lavender Mist";

    t.menuBarBackground       = QColor("#5C3D8F");
    t.menuBarHover            = QColor("#4A3175");
    t.statusBarBackground     = QColor("#7E5DAF");
    t.toolbarAccentStripe     = QColor("#5C3D8F");
    t.tabActiveIndicator      = QColor("#7E5DAF");
    t.focusBorderColor        = QColor("#7E5DAF");
    t.editorBorderColor       = QColor("#7E5DAF");
    t.windowBackground        = QColor("#F2EFF5");

    t.chatHeaderGradientStart = QColor("#5C3D8F");
    t.chatHeaderGradientEnd   = QColor("#8E6CBB");
    t.chatBackground          = QColor("#F0EAF5");
    t.chatUserBubble          = QColor("#E8DFF2");
    t.chatSendButton          = QColor("#A78BCC");
    t.chatSendButtonHover     = QColor("#8E6CBB");

    t.dockTitleBackground     = QColor("#5C3D8F");

    t.dialogButtonPrimary     = QColor("#7E5DAF");
    t.dialogButtonPrimaryHover = QColor("#5C3D8F");

    t.accentPrimary  = QColor("#A78BCC");
    t.accentDark     = QColor("#7E5DAF");
    t.accentDarker   = QColor("#5C3D8F");
    t.accentLight    = QColor("#E8DFF2");

    t.checkboxChecked         = QColor("#A78BCC");
    t.selectionHandleColor    = QColor("#7E5DAF");

    t.menuItemHover           = QColor("#EDE6F5");
    t.toolbarButtonChecked    = QColor("#E0D4F0");
    t.popupItemSelected       = QColor("#EDE6F5");

    return t;
}

static NexelTheme createArcticFrost() {
    NexelTheme t = createNexelGreen();
    t.id = "arctic_frost";
    t.displayName = "Arctic Frost";

    t.menuBarBackground       = QColor("#3C5E72");
    t.menuBarHover            = QColor("#304D5E");
    t.statusBarBackground     = QColor("#527D96");
    t.toolbarAccentStripe     = QColor("#3C5E72");
    t.tabActiveIndicator      = QColor("#527D96");
    t.focusBorderColor        = QColor("#527D96");
    t.editorBorderColor       = QColor("#527D96");
    t.windowBackground        = QColor("#EDF0F2");

    t.chatHeaderGradientStart = QColor("#3C5E72");
    t.chatHeaderGradientEnd   = QColor("#6899AC");
    t.chatBackground          = QColor("#E6ECF0");
    t.chatUserBubble          = QColor("#DAE8EF");
    t.chatSendButton          = QColor("#7BA7BC");
    t.chatSendButtonHover     = QColor("#6899AC");

    t.dockTitleBackground     = QColor("#3C5E72");

    t.dialogButtonPrimary     = QColor("#527D96");
    t.dialogButtonPrimaryHover = QColor("#3C5E72");

    t.accentPrimary  = QColor("#7BA7BC");
    t.accentDark     = QColor("#527D96");
    t.accentDarker   = QColor("#3C5E72");
    t.accentLight    = QColor("#DAE8EF");

    t.checkboxChecked         = QColor("#7BA7BC");
    t.selectionHandleColor    = QColor("#527D96");

    t.menuItemHover           = QColor("#E0EAF0");
    t.toolbarButtonChecked    = QColor("#D0DEE8");
    t.popupItemSelected       = QColor("#E0EAF0");

    return t;
}

static NexelTheme createSunsetGlow() {
    NexelTheme t = createNexelGreen();
    t.id = "sunset_glow";
    t.displayName = "Sunset Glow";

    t.menuBarBackground       = QColor("#A0513A");
    t.menuBarHover            = QColor("#874430");
    t.statusBarBackground     = QColor("#C96F4A");
    t.toolbarAccentStripe     = QColor("#A0513A");
    t.tabActiveIndicator      = QColor("#C96F4A");
    t.focusBorderColor        = QColor("#C96F4A");
    t.editorBorderColor       = QColor("#C96F4A");
    t.windowBackground        = QColor("#F5F0ED");

    t.chatHeaderGradientStart = QColor("#A0513A");
    t.chatHeaderGradientEnd   = QColor("#D08060");
    t.chatBackground          = QColor("#F5EDE8");
    t.chatUserBubble          = QColor("#FAEADE");
    t.chatSendButton          = QColor("#E8916D");
    t.chatSendButtonHover     = QColor("#D08060");

    t.dockTitleBackground     = QColor("#A0513A");

    t.dialogButtonPrimary     = QColor("#C96F4A");
    t.dialogButtonPrimaryHover = QColor("#A0513A");

    t.accentPrimary  = QColor("#E8916D");
    t.accentDark     = QColor("#C96F4A");
    t.accentDarker   = QColor("#A0513A");
    t.accentLight    = QColor("#FAEADE");

    t.checkboxChecked         = QColor("#E8916D");
    t.selectionHandleColor    = QColor("#C96F4A");

    t.menuItemHover           = QColor("#F5E8E0");
    t.toolbarButtonChecked    = QColor("#F0D8C8");
    t.popupItemSelected       = QColor("#F5E8E0");

    return t;
}

// ── ThemeManager Implementation ──

ThemeManager& ThemeManager::instance() {
    static ThemeManager mgr;
    return mgr;
}

ThemeManager::ThemeManager() {
    registerThemes();
}

void ThemeManager::registerThemes() {
    m_themes = {
        createNexelGreen(),
        createOceanBreeze(),
        createRoseQuartz(),
        createLavenderMist(),
        createArcticFrost(),
        createSunsetGlow()
    };
}

QVector<NexelTheme> ThemeManager::availableThemes() const {
    return m_themes;
}

const NexelTheme& ThemeManager::currentTheme() const {
    return m_themes[m_currentIndex];
}

void ThemeManager::loadSavedTheme() {
    QSettings settings("Nexel", "Nexel");
    QString savedId = settings.value("theme", "nexel_green").toString();
    for (int i = 0; i < m_themes.size(); ++i) {
        if (m_themes[i].id == savedId) {
            m_currentIndex = i;
            return;
        }
    }
    m_currentIndex = 0;
}

void ThemeManager::setTheme(const QString& themeId) {
    for (int i = 0; i < m_themes.size(); ++i) {
        if (m_themes[i].id == themeId) {
            m_currentIndex = i;
            QSettings settings("Nexel", "Nexel");
            settings.setValue("theme", themeId);
            emit themeChanged(m_themes[i]);
            return;
        }
    }
}

void ThemeManager::applyTheme(QMainWindow* window) {
    window->setStyleSheet(buildGlobalStylesheet());
}

QString ThemeManager::buildGlobalStylesheet() const {
    const auto& t = currentTheme();
    return QString(
        "QMainWindow { background-color: %1; }"
        "QMenuBar { background-color: %2; color: %3; border: none; padding: 2px; font-size: 12px; }"
        "QMenuBar::item { padding: 4px 12px; border-radius: 3px; }"
        "QMenuBar::item:selected { background-color: %4; }"
        "QMenu { background-color: %5; border: 1px solid %6; border-radius: 6px; padding: 4px; }"
        "QMenu::item { padding: 6px 30px 6px 20px; border-radius: 4px; }"
        "QMenu::item:selected { background-color: %7; }"
        "QMenu::separator { height: 1px; background: %8; margin: 4px 8px; }"
        "QStatusBar { background-color: %9; color: %10; border: none; font-size: 11px; padding: 2px 8px; }"
    ).arg(
        t.windowBackground.name(),
        t.menuBarBackground.name(),
        t.menuBarText.name(),
        t.menuBarHover.name(),
        t.menuBackground.name(),
        t.menuBorder.name(),
        t.menuItemHover.name(),
        t.menuSeparator.name(),
        t.statusBarBackground.name()
    ).arg(
        t.statusBarText.name()
    );
}

QString ThemeManager::buildTabBarStylesheet() const {
    const auto& t = currentTheme();
    return QString(
        "QTabBar { background: transparent; border: none; }"
        "QTabBar::tab {"
        "   background-color: %1;"
        "   border: 1px solid %2;"
        "   border-bottom: none;"
        "   padding: 3px 16px;"
        "   margin-right: 2px;"
        "   font-size: 11px;"
        "   min-width: 60px;"
        "   border-top-left-radius: 3px;"
        "   border-top-right-radius: 3px;"
        "}"
        "QTabBar::tab:selected {"
        "   background-color: %3;"
        "   border-bottom: 2px solid %4;"
        "   font-weight: bold;"
        "}"
        "QTabBar::tab:hover:!selected {"
        "   background-color: %5;"
        "}"
    ).arg(
        t.tabBackground.name(),
        t.tabBorder.name(),
        t.tabActiveBackground.name(),
        t.tabActiveIndicator.name(),
        t.tabHover.name()
    );
}

QString ThemeManager::buildBottomBarStylesheet() const {
    const auto& t = currentTheme();
    return QString("QWidget { background-color: %1; border-top: 1px solid %2; }")
        .arg(t.bottomBarBackground.name(), t.bottomBarBorder.name());
}

QString ThemeManager::buildAddSheetBtnStylesheet() const {
    const auto& t = currentTheme();
    return QString(
        "QToolButton { background: transparent; border: 1px solid transparent; "
        "border-radius: 3px; font-size: 16px; font-weight: bold; color: %1; }"
        "QToolButton:hover { background-color: %2; border-color: #C0C0C0; }"
    ).arg(t.addSheetButtonText.name(), t.addSheetButtonHover.name());
}

QString ThemeManager::dialogStylesheet() {
    const auto& t = instance().currentTheme();
    return QString(
        "QDialog { background: %1; }"
        "QGroupBox { font-weight: 600; font-size: 12px; border: 1px solid %2; "
        "border-radius: 10px; margin-top: 10px; padding: 20px 14px 14px 14px; background: white; }"
        "QGroupBox::title { subcontrol-origin: margin; left: 14px; padding: 0 8px; "
        "color: %3; font-size: 12px; }"
        "QLineEdit { border: 1px solid %2; border-radius: 6px; padding: 6px 10px; "
        "background: white; font-size: 12px; color: #1D2939; }"
        "QLineEdit:focus { border-color: %4; }"
        "QLineEdit::placeholder { color: #98A2B3; }"
        "QComboBox { border: 1px solid %2; border-radius: 6px; padding: 6px 10px; "
        "background: white; font-size: 12px; color: #1D2939; min-height: 22px; }"
        "QComboBox:focus { border: 1px solid %4; }"
        "QComboBox::drop-down { border: none; width: 22px; }"
        "QComboBox::down-arrow { image: none; border-left: 4px solid transparent; "
        "border-right: 4px solid transparent; border-top: 5px solid #667085; margin-right: 6px; }"
        "QComboBox QAbstractItemView { border: 1px solid %2; border-radius: 6px; "
        "background: white; selection-background-color: %5; padding: 4px; outline: none; }"
        "QCheckBox { spacing: 8px; font-size: 12px; color: #344054; }"
        "QListWidget { border: 1px solid %2; border-radius: 8px; background: white; color: #1D2939; outline: none; }"
        "QListWidget::item { padding: 8px 10px; border-radius: 6px; font-size: 12px; color: #1D2939; }"
        "QListWidget::item:selected { background-color: %5; color: #1D2939; "
        "border-left: 3px solid %3; }"
        "QListWidget::item:hover:!selected { background-color: %1; }"
        "QLabel { color: #475467; font-size: 12px; }"
    ).arg(
        t.dialogBackground.name(),
        t.dialogGroupBoxBorder.name(),
        t.accentDark.name(),
        t.accentPrimary.name(),
        t.accentLight.name()
    );
}
