#include "ZipStreamReader.h"
#include "mz.h"
#include "mz_strm.h"
#include "mz_zip.h"
#include "mz_zip_rw.h"

#include <cstring>

ZipStreamReader::ZipStreamReader(const QString& filePath) {
    m_handle = mz_zip_reader_create();
    if (!m_handle) {
        m_lastError = "mz_zip_reader_create failed";
        return;
    }
    const QByteArray pathUtf8 = filePath.toUtf8();
    int32_t rc = mz_zip_reader_open_file(m_handle, pathUtf8.constData());
    if (rc != MZ_OK) {
        m_lastError = QString("mz_zip_reader_open_file rc=%1").arg(rc);
        mz_zip_reader_delete(&m_handle);
        m_handle = nullptr;
        return;
    }
    m_open = true;
}

ZipStreamReader::~ZipStreamReader() {
    closeActiveEntry();
    if (m_open) {
        mz_zip_reader_close(m_handle);
        m_open = false;
    }
    if (m_handle) {
        mz_zip_reader_delete(&m_handle);
        m_handle = nullptr;
    }
}

void ZipStreamReader::closeActiveEntry() {
    if (m_entryOpen) {
        mz_zip_reader_entry_close(m_handle);
        m_entryOpen = false;
    }
}

QStringList ZipStreamReader::entries() const {
    QStringList out;
    if (!m_open) return out;

    int32_t rc = mz_zip_reader_goto_first_entry(m_handle);
    while (rc == MZ_OK) {
        mz_zip_file* info = nullptr;
        if (mz_zip_reader_entry_get_info(m_handle, &info) == MZ_OK && info && info->filename) {
            out.append(QString::fromUtf8(info->filename));
        }
        rc = mz_zip_reader_goto_next_entry(m_handle);
    }
    return out;
}

qint64 ZipStreamReader::entrySize(const QString& path) const {
    if (!m_open) return -1;
    const QByteArray nameUtf8 = path.toUtf8();
    if (mz_zip_reader_locate_entry(m_handle, nameUtf8.constData(), 0) != MZ_OK) {
        return -1;
    }
    mz_zip_file* info = nullptr;
    if (mz_zip_reader_entry_get_info(m_handle, &info) != MZ_OK || !info) {
        return -1;
    }
    return static_cast<qint64>(info->uncompressed_size);
}

std::unique_ptr<ZipEntryDevice> ZipStreamReader::openEntry(const QString& path) {
    if (!m_open) return nullptr;

    closeActiveEntry();

    const QByteArray nameUtf8 = path.toUtf8();
    if (mz_zip_reader_locate_entry(m_handle, nameUtf8.constData(), 0) != MZ_OK) {
        m_lastError = QString("entry not found: %1").arg(path);
        return nullptr;
    }

    mz_zip_file* info = nullptr;
    if (mz_zip_reader_entry_get_info(m_handle, &info) != MZ_OK || !info) {
        m_lastError = "mz_zip_reader_entry_get_info failed";
        return nullptr;
    }
    const qint64 uncompressed = static_cast<qint64>(info->uncompressed_size);

    if (mz_zip_reader_entry_open(m_handle) != MZ_OK) {
        m_lastError = QString("mz_zip_reader_entry_open failed for %1").arg(path);
        return nullptr;
    }
    m_entryOpen = true;

    std::unique_ptr<ZipEntryDevice> dev(new ZipEntryDevice(this, uncompressed));
    if (!dev->open(QIODevice::ReadOnly)) {
        m_lastError = "ZipEntryDevice::open failed";
        closeActiveEntry();
        return nullptr;
    }
    return dev;
}

QByteArray ZipStreamReader::readEntry(const QString& path) {
    QByteArray out;
    auto dev = openEntry(path);
    if (!dev) return out;
    constexpr qint64 CHUNK = 64 * 1024;
    QByteArray buf;
    buf.resize(CHUNK);
    qint64 n;
    while ((n = dev->read(buf.data(), CHUNK)) > 0) {
        out.append(buf.constData(), static_cast<int>(n));
    }
    return out;
}

// ----- ZipEntryDevice -----

ZipEntryDevice::ZipEntryDevice(ZipStreamReader* owner, qint64 uncompressedSize)
    : m_owner(owner), m_uncompressedSize(uncompressedSize) {}

ZipEntryDevice::~ZipEntryDevice() {
    if (m_owner) {
        m_owner->closeActiveEntry();
    }
}

qint64 ZipEntryDevice::bytesAvailable() const {
    return (m_uncompressedSize - m_bytesRead) + QIODevice::bytesAvailable();
}

qint64 ZipEntryDevice::readData(char* data, qint64 maxlen) {
    if (!m_owner || !m_owner->m_entryOpen) return -1;
    if (maxlen <= 0) return 0;
    if (m_bytesRead >= m_uncompressedSize) return 0;

    const qint64 remaining = m_uncompressedSize - m_bytesRead;
    const qint64 wanted = qMin(maxlen, remaining);

    const int32_t n = mz_zip_reader_entry_read(m_owner->m_handle, data, static_cast<int32_t>(wanted));
    if (n < 0) return -1;
    m_bytesRead += n;
    return n;
}
