#include "AutoSaveService.h"
#include "../ui/MainWindow.h"
#include "../core/Spreadsheet.h"
#include "../core/ColumnStore.h"
#include "../core/StyleTable.h"
#include "XlsxService.h"

#include <QCryptographicHash>
#include <QDir>
#include <QFile>
#include <QFileInfo>
#include <QStandardPaths>
#include <QDebug>
#include <QtConcurrent>

AutoSaveService::AutoSaveService(QObject* parent)
    : QObject(parent) {
    m_timer.setTimerType(Qt::VeryCoarseTimer);
    connect(&m_timer, &QTimer::timeout, this, &AutoSaveService::performAutoSave);
}

AutoSaveService::~AutoSaveService() {
    m_timer.stop();
    // Wait for any in-progress save to finish
    if (m_saveThread) {
        m_saveThread->quit();
        m_saveThread->wait(5000);
        delete m_saveThread;
    }
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

QString AutoSaveService::checkForRecovery(const QString& originalFilePath) {
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

void AutoSaveService::cleanupAutoSave(const QString& originalFilePath) {
    QString autoSavePath = getAutoSavePath(originalFilePath);
    if (QFile::exists(autoSavePath)) {
        QFile::remove(autoSavePath);
    }
    if (originalFilePath.isEmpty()) return;
    QString unsavedPath = getAutoSavePath(QString());
    if (QFile::exists(unsavedPath)) {
        QFile::remove(unsavedPath);
    }
}

void AutoSaveService::onManualSave() {
    m_isDirty = false;
    cleanupAutoSave(m_currentFilePath);
    if (m_enabled) {
        m_timer.start(m_intervalSeconds * 1000);
    }
}

void AutoSaveService::setCurrentFilePath(const QString& path) {
    if (!m_currentFilePath.isEmpty()) {
        cleanupAutoSave(m_currentFilePath);
    }
    m_currentFilePath = path;
}

void AutoSaveService::markDirty() {
    m_isDirty = true;
    if (m_enabled && !m_timer.isActive()) {
        m_timer.start(m_intervalSeconds * 1000);
    }
}

void AutoSaveService::performAutoSave() {
    if (!m_isDirty || !m_mainWindow) return;

    // Don't start another save if one is already running
    if (m_saving.load()) {
        qDebug() << "AutoSave: previous save still in progress, skipping";
        return;
    }

    const auto& sheets = m_mainWindow->getSheets();
    if (sheets.empty()) return;

    // Check if any sheet is large enough to need background save
    bool isLarge = false;
    for (const auto& sheet : sheets) {
        if (sheet && sheet->getRowCount() > 100000) {
            isLarge = true;
            break;
        }
    }

    QString autoSavePath = getAutoSavePath(m_currentFilePath);

    if (!isLarge) {
        // Small files: save synchronously (fast, <100ms)
        std::vector<NexelChartExport> emptyCharts;
        bool success = XlsxService::exportToFile(sheets, autoSavePath, emptyCharts);
        if (success) {
            m_isDirty = false;
            qDebug() << "AutoSave: saved to" << autoSavePath;
        }
        return;
    }

    // Large files: save in background thread (zero UI blocking)
    m_saving.store(true);
    emit autoSaveStarted();

    // Take a lightweight snapshot of sheet data on the main thread
    // This is fast: just copies the shared_ptr (reference counted)
    std::vector<std::shared_ptr<Spreadsheet>> sheetsCopy(sheets.begin(), sheets.end());
    QString savePath = autoSavePath;

    QtConcurrent::run([this, sheetsCopy = std::move(sheetsCopy), savePath]() {
        std::vector<NexelChartExport> emptyCharts;
        bool success = XlsxService::exportToFile(sheetsCopy, savePath, emptyCharts);

        // Signal completion back to main thread
        QMetaObject::invokeMethod(this, [this, success]() {
            m_saving.store(false);
            if (success) {
                m_isDirty = false;
                qDebug() << "AutoSave: background save completed";
            } else {
                qWarning() << "AutoSave: background save failed";
            }
            emit autoSaveFinished(success);
        }, Qt::QueuedConnection);
    });
}
