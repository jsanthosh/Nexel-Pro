#include <QApplication>
#include <QStandardPaths>
#include <QDir>
#include "ui/MainWindow.h"
#include "ui/Theme.h"
#include "database/DatabaseManager.h"
#include "services/DocumentService.h"

int main(int argc, char* argv[]) {
    QApplication app(argc, argv);

    app.setApplicationName("Nexel");
    app.setApplicationVersion("4.1.0");
    app.setApplicationDisplayName("Nexel");

    // Initialize database
    QString appDataPath = QStandardPaths::writableLocation(QStandardPaths::AppDataLocation);
    QDir().mkpath(appDataPath);
    QString dbPath = appDataPath + "/documents.db";

    if (!DatabaseManager::instance().initialize(dbPath)) {
        qWarning() << "Failed to initialize database:" << DatabaseManager::instance().getLastError();
        return 1;
    }

    // Load saved theme
    ThemeManager::instance().loadSavedTheme();

    // Create main window
    MainWindow window;
    ThemeManager::instance().applyTheme(&window);
    window.show();

    // Open file passed as command-line argument
    if (argc > 1) {
        window.openFile(QString::fromLocal8Bit(argv[1]));
    }

    return app.exec();
}
