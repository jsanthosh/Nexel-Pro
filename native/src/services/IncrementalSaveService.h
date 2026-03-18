#ifndef INCREMENTALSAVESERVICE_H
#define INCREMENTALSAVESERVICE_H

#include <QObject>
#include <QTimer>
#include <QThread>
#include <QMutex>
#include <QString>
#include <QVariant>
#include <vector>
#include <memory>
#include <atomic>

class MainWindow;
class Spreadsheet;

// ============================================================================
// IncrementalSaveService — Background save with delta journal + COW snapshots
// ============================================================================
// Replaces the synchronous AutoSaveService for 20M-row workbooks.
//
// Small edits: flushed to an append-only SQLite journal (<1ms).
// Large changes: COW snapshot of ColumnStore chunks → background XLSX export.
// The UI thread is NEVER blocked during save.
//
class IncrementalSaveService : public QObject {
    Q_OBJECT

public:
    explicit IncrementalSaveService(QObject* parent = nullptr);
    ~IncrementalSaveService();

    void setMainWindow(MainWindow* mainWindow);
    void setInterval(int seconds);
    void setEnabled(bool enabled);
    bool isEnabled() const;

    // Called when user saves normally
    void onManualSave();
    void setCurrentFilePath(const QString& path);

    // Check for recovery (delegates to AutoSaveService-compatible paths)
    static QString checkForRecovery(const QString& originalFilePath);
    static QString getAutoSavePath(const QString& originalFilePath);
    static void cleanupAutoSave(const QString& originalFilePath);

    // Record a cell change in the delta journal
    void recordCellChange(int sheetIndex, int row, int col, const QVariant& oldValue, const QVariant& newValue);

    // Is a background save currently running?
    bool isSaving() const { return m_isSaving.load(); }

signals:
    void saveStarted();
    void saveCompleted(bool success);
    void saveProgress(int percent);

public slots:
    void markDirty();

private slots:
    void performAutoSave();
    void onBackgroundSaveFinished();

private:
    MainWindow* m_mainWindow = nullptr;
    QTimer m_timer;
    QString m_currentFilePath;
    bool m_enabled = true;
    int m_intervalSeconds = 60;
    bool m_isDirty = false;
    std::atomic<bool> m_isSaving{false};

    // Delta journal: lightweight record of cell changes since last save
    struct CellDelta {
        int sheetIndex;
        int row;
        int col;
        QVariant newValue;
    };
    QMutex m_journalMutex;
    std::vector<CellDelta> m_journal;

    // Background save thread
    QThread m_saveThread;

    // Threshold: if journal exceeds this, do a full background save
    // Otherwise, just flush the journal to the recovery file
    static constexpr int FULL_SAVE_THRESHOLD = 10000;

    void flushJournal();
    void triggerBackgroundFullSave();
};

#endif // INCREMENTALSAVESERVICE_H
