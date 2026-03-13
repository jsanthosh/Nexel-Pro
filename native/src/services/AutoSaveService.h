#ifndef AUTOSAVESERVICE_H
#define AUTOSAVESERVICE_H

#include <QObject>
#include <QTimer>
#include <QString>

class MainWindow;

class AutoSaveService : public QObject {
    Q_OBJECT
public:
    explicit AutoSaveService(QObject* parent = nullptr);
    ~AutoSaveService();

    void setMainWindow(MainWindow* mainWindow);
    void setInterval(int seconds);  // default 60
    void setEnabled(bool enabled);
    bool isEnabled() const;

    // Check for orphaned auto-save files on startup
    // Returns the path of a recovery file if found, empty string otherwise
    static QString checkForRecovery(const QString& originalFilePath);
    static QString getAutoSavePath(const QString& originalFilePath);
    static void cleanupAutoSave(const QString& originalFilePath);

    // Called when user saves normally — reset timer and clean auto-save
    void onManualSave();
    // Called when file path changes (Save As)
    void setCurrentFilePath(const QString& path);

private slots:
    void performAutoSave();

private:
    MainWindow* m_mainWindow = nullptr;
    QTimer m_timer;
    QString m_currentFilePath;
    bool m_enabled = true;
    int m_intervalSeconds = 60;
    bool m_isDirty = false;

public slots:
    void markDirty();  // Call when any cell changes
};

#endif
