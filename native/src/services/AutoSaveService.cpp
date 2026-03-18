#include "AutoSaveService.h"
#include "../ui/MainWindow.h"
#include "XlsxService.h"

#include <QCryptographicHash>
#include <QDir>
#include <QFile>
#include <QFileInfo>
#include <QStandardPaths>
#include <QDebug>

AutoSaveService::AutoSaveService(QObject* parent)
    : QObject(parent) {
    m_timer.setTimerType(Qt::VeryCoarseTimer);
    connect(&m_timer, &QTimer::timeout, this, &AutoSaveService::performAutoSave);
}

AutoSaveService::~AutoSaveService() {
    m_timer.stop();
}

void AutoSaveService::setMainWindow(MainWindow* mainWindow) {
    m_mainWindow = mainWindow;
}

void AutoSaveService::setInterval(int seconds) {
    m_intervalSeconds = seconds;
    if (m_timer.isActive()) {
        m_timer.start(m_intervalSeconds * 1000);
    }
}

void AutoSaveService::setEnabled(bool enabled) {
    m_enabled = enabled;
    if (m_enabled) {
        m_timer.start(m_intervalSeconds * 1000);
    } else {
        m_timer.stop();
    }
}

bool AutoSaveService::isEnabled() const {
    return m_enabled;
}

QString AutoSaveService::getAutoSavePath(const QString& originalFilePath) {
    if (originalFilePath.isEmpty()) {
        // Unsaved document — use temp directory
        QString tempDir = QStandardPaths::writableLocation(QStandardPaths::TempLocation);
        return QDir(tempDir).filePath(".~nexel-unsaved-recovery.xlsx");
    }

    // Generate hex hash of original path for uniqueness
    QByteArray hash = QCryptographicHash::hash(
        originalFilePath.toUtf8(), QCryptographicHash::Md5).toHex();
    QString hashStr = QString::fromLatin1(hash.left(16));  // Use first 16 hex chars

    QFileInfo fi(originalFilePath);
    QString dir = fi.absolutePath();
    return QDir(dir).filePath(".~nexel-autosave-" + hashStr + ".xlsx");
}

QString AutoSaveService::checkForRecovery(const QString& originalFilePath) {
    QString autoSavePath = getAutoSavePath(originalFilePath);
    QFileInfo autoSaveInfo(autoSavePath);

    if (!autoSaveInfo.exists()) {
        return QString();
    }

    if (originalFilePath.isEmpty()) {
        // Unsaved document — any auto-save file is a recovery candidate
        return autoSavePath;
    }

    QFileInfo originalInfo(originalFilePath);
    if (!originalInfo.exists()) {
        // Original file doesn't exist but auto-save does — recovery candidate
        return autoSavePath;
    }

    // Only offer recovery if auto-save is newer than the original
    if (autoSaveInfo.lastModified() > originalInfo.lastModified()) {
        return autoSavePath;
    }

    // Auto-save is older, clean it up
    QFile::remove(autoSavePath);
    return QString();
}

void AutoSaveService::cleanupAutoSave(const QString& originalFilePath) {
    QString autoSavePath = getAutoSavePath(originalFilePath);
    if (QFile::exists(autoSavePath)) {
        QFile::remove(autoSavePath);
    }

    // Also clean up the unsaved document recovery file
    if (originalFilePath.isEmpty()) {
        return;
    }
    // If original path is set, also try to clean up the unsaved recovery
    QString unsavedPath = getAutoSavePath(QString());
    if (QFile::exists(unsavedPath)) {
        QFile::remove(unsavedPath);
    }
}

void AutoSaveService::onManualSave() {
    m_isDirty = false;
    cleanupAutoSave(m_currentFilePath);
    // Restart timer from scratch after manual save
    if (m_enabled) {
        m_timer.start(m_intervalSeconds * 1000);
    }
}

void AutoSaveService::setCurrentFilePath(const QString& path) {
    // Clean up auto-save for old path
    if (!m_currentFilePath.isEmpty()) {
        cleanupAutoSave(m_currentFilePath);
    }
    m_currentFilePath = path;
}

void AutoSaveService::markDirty() {
    m_isDirty = true;
    // Start the timer if not already running
    if (m_enabled && !m_timer.isActive()) {
        m_timer.start(m_intervalSeconds * 1000);
    }
}

void AutoSaveService::performAutoSave() {
    if (!m_isDirty || !m_mainWindow) {
        return;
    }

    // Skip auto-save for large datasets (>500K rows) — synchronous XLSX export
    // would freeze the UI for tens of seconds. TODO: implement async background save.
    const auto& sheets = m_mainWindow->getSheets();
    if (!sheets.empty()) {
        for (const auto& sheet : sheets) {
            if (sheet && sheet->getRowCount() > 500000) {
                qDebug() << "AutoSave: skipped (large dataset:" << sheet->getRowCount() << "rows)";
                return;
            }
        }
    }

    QString autoSavePath = getAutoSavePath(m_currentFilePath);

    if (sheets.empty()) {
        qWarning() << "AutoSave: No sheets to save";
        return;
    }

    std::vector<NexelChartExport> emptyCharts;
    bool success = XlsxService::exportToFile(sheets, autoSavePath, emptyCharts);

    if (success) {
        m_isDirty = false;
        qDebug() << "AutoSave: saved to" << autoSavePath;
    } else {
        qWarning() << "AutoSave: failed to save to" << autoSavePath;
    }
}
