#ifndef AUTOSAVESERVICE_H
#define AUTOSAVESERVICE_H

#include <QObject>
#include <QTimer>
#include <QString>
#include <QThread>
#include <QMutex>
#include <atomic>

class MainWindow;
class Spreadsheet;

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
    static QString checkForRecovery(const QString& originalFilePath);
    static QString getAutoSavePath(const QString& originalFilePath);
    static void cleanupAutoSave(const QString& originalFilePath);

    // Called when user saves normally — reset timer and clean auto-save
    void onManualSave();
    // Called when file path changes (Save As)
    void setCurrentFilePath(const QString& path);

    // Check if background save is in progress
    bool isSaving() const { return m_saving.load(); }

signals:
    void autoSaveStarted();
    void autoSaveFinished(bool success);

private slots:
    void performAutoSave();

private:
    MainWindow* m_mainWindow = nullptr;
    QTimer m_timer;
    QString m_currentFilePath;
    bool m_enabled = true;
    int m_intervalSeconds = 60;
    bool m_isDirty = false;
    std::atomic<bool> m_saving{false};

    // Background save thread
    QThread* m_saveThread = nullptr;

    // Snapshot cell data for background save (COW-style)
    struct SheetSnapshot {
        QString name;
        int rowCount, colCount;
        std::vector<std::tuple<int, int, QVariant, uint16_t, QString>> cells;
        // (row, col, value, styleIdx, formula)
    };
    std::vector<SheetSnapshot> snapshotSheets();

public slots:
    void markDirty();
};

#endif
