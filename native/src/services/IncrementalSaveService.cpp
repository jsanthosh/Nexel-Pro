#include "IncrementalSaveService.h"
#include "../ui/MainWindow.h"
#include "XlsxService.h"

#include <QCryptographicHash>
#include <QDir>
#include <QFile>
#include <QFileInfo>
#include <QStandardPaths>
#include <QDebug>
#include <QtConcurrent>

IncrementalSaveService::IncrementalSaveService(QObject* parent)
    : QObject(parent) {
    m_timer.setTimerType(Qt::VeryCoarseTimer);
    connect(&m_timer, &QTimer::timeout, this, &IncrementalSaveService::performAutoSave);
}

IncrementalSaveService::~IncrementalSaveService() {
    m_timer.stop();
    if (m_saveThread.isRunning()) {
        m_saveThread.quit();
        m_saveThread.wait(5000);
    }
}

void IncrementalSaveService::setMainWindow(MainWindow* mainWindow) {
    m_mainWindow = mainWindow;
}

void IncrementalSaveService::setInterval(int seconds) {
    m_intervalSeconds = seconds;
    if (m_timer.isActive()) {
        m_timer.start(m_intervalSeconds * 1000);
    }
}

void IncrementalSaveService::setEnabled(bool enabled) {
    m_enabled = enabled;
    if (m_enabled) {
        m_timer.start(m_intervalSeconds * 1000);
    } else {
        m_timer.stop();
    }
}

bool IncrementalSaveService::isEnabled() const {
    return m_enabled;
}

QString IncrementalSaveService::getAutoSavePath(const QString& originalFilePath) {
    if (originalFilePath.isEmpty()) {
        QString tempDir = QStandardPaths::writableLocation(QStandardPaths::TempLocation);
        return QDir(tempDir).filePath(".~nexel-unsaved-recovery.xlsx");
    }

    QByteArray hash = QCryptographicHash::hash(
        originalFilePath.toUtf8(), QCryptographicHash::Md5).toHex();
    QString hashStr = QString::fromLatin1(hash.left(16));

    QFileInfo fi(originalFilePath);
    QString dir = fi.absolutePath();
    return QDir(dir).filePath(".~nexel-autosave-" + hashStr + ".xlsx");
}

QString IncrementalSaveService::checkForRecovery(const QString& originalFilePath) {
    QString autoSavePath = getAutoSavePath(originalFilePath);
    QFileInfo autoSaveInfo(autoSavePath);

    if (!autoSaveInfo.exists()) return QString();

    if (originalFilePath.isEmpty()) return autoSavePath;

    QFileInfo originalInfo(originalFilePath);
    if (!originalInfo.exists()) return autoSavePath;

    if (autoSaveInfo.lastModified() > originalInfo.lastModified()) {
        return autoSavePath;
    }

    QFile::remove(autoSavePath);
    return QString();
}

void IncrementalSaveService::cleanupAutoSave(const QString& originalFilePath) {
    QString autoSavePath = getAutoSavePath(originalFilePath);
    if (QFile::exists(autoSavePath)) {
        QFile::remove(autoSavePath);
    }
    if (!originalFilePath.isEmpty()) {
        QString unsavedPath = getAutoSavePath(QString());
        if (QFile::exists(unsavedPath)) {
            QFile::remove(unsavedPath);
        }
    }
}

void IncrementalSaveService::onManualSave() {
    m_isDirty = false;
    {
        QMutexLocker lock(&m_journalMutex);
        m_journal.clear();
    }
    cleanupAutoSave(m_currentFilePath);
    if (m_enabled) {
        m_timer.start(m_intervalSeconds * 1000);
    }
}

void IncrementalSaveService::setCurrentFilePath(const QString& path) {
    if (!m_currentFilePath.isEmpty()) {
        cleanupAutoSave(m_currentFilePath);
    }
    m_currentFilePath = path;
}

void IncrementalSaveService::markDirty() {
    m_isDirty = true;
    if (m_enabled && !m_timer.isActive()) {
        m_timer.start(m_intervalSeconds * 1000);
    }
}

void IncrementalSaveService::recordCellChange(int sheetIndex, int row, int col,
                                                const QVariant& /*oldValue*/, const QVariant& newValue) {
    QMutexLocker lock(&m_journalMutex);
    m_journal.push_back({sheetIndex, row, col, newValue});
    m_isDirty = true;
}

void IncrementalSaveService::performAutoSave() {
    if (!m_isDirty || !m_mainWindow) return;
    if (m_isSaving.load()) return; // Already saving

    int journalSize;
    {
        QMutexLocker lock(&m_journalMutex);
        journalSize = static_cast<int>(m_journal.size());
    }

    if (journalSize < FULL_SAVE_THRESHOLD && journalSize > 0) {
        // Small number of changes: quick journal flush
        flushJournal();
    } else {
        // Large changes or no journal: full background save
        triggerBackgroundFullSave();
    }
}

void IncrementalSaveService::flushJournal() {
    // For journal flush, we still do a full XLSX save but it's acceptable
    // for small change counts. The key improvement is that for large workbooks,
    // triggerBackgroundFullSave() runs on a background thread.
    triggerBackgroundFullSave();
}

void IncrementalSaveService::triggerBackgroundFullSave() {
    if (m_isSaving.exchange(true)) return; // Already saving

    emit saveStarted();

    // Capture sheets snapshot (shared_ptr copies are thread-safe)
    const auto& sheets = m_mainWindow->getSheets();
    if (sheets.empty()) {
        m_isSaving.store(false);
        return;
    }

    // Copy shared_ptrs (COW: if UI modifies a chunk during save, it copies that chunk)
    auto sheetsCopy = sheets;
    QString savePath = getAutoSavePath(m_currentFilePath);

    // Run the XLSX export on a background thread via QtConcurrent
    auto future = QtConcurrent::run([sheetsCopy, savePath]() -> bool {
        std::vector<NexelChartExport> emptyCharts;
        return XlsxService::exportToFile(sheetsCopy, savePath, emptyCharts);
    });

    // Use a watcher to get notified when done (back on main thread)
    auto* watcher = new QFutureWatcher<bool>(this);
    connect(watcher, &QFutureWatcher<bool>::finished, this, [this, watcher]() {
        bool success = watcher->result();
        watcher->deleteLater();

        m_isSaving.store(false);
        if (success) {
            m_isDirty = false;
            {
                QMutexLocker lock(&m_journalMutex);
                m_journal.clear();
            }
            qDebug() << "IncrementalSave: background save completed successfully";
        } else {
            qWarning() << "IncrementalSave: background save failed";
        }
        emit saveCompleted(success);
    });
    watcher->setFuture(future);
}

void IncrementalSaveService::onBackgroundSaveFinished() {
    // Handled by the lambda in triggerBackgroundFullSave
}
