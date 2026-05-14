#include "ZipStreamWriter.h"
#include "mz.h"
#include "mz_strm.h"
#include "mz_zip.h"
#include "mz_zip_rw.h"

#include <cstring>

ZipStreamWriter::ZipStreamWriter(const QString& filePath) {
    m_handle = mz_zip_writer_create();
    if (!m_handle) {
        m_lastError = "mz_zip_writer_create failed";
        return;
    }
    mz_zip_writer_set_compress_method(m_handle, MZ_COMPRESS_METHOD_DEFLATE);
    mz_zip_writer_set_compress_level(m_handle, MZ_COMPRESS_LEVEL_DEFAULT);

    const QByteArray pathUtf8 = filePath.toUtf8();
    const int32_t rc = mz_zip_writer_open_file(m_handle, pathUtf8.constData(), 0 /*disk_size*/, 0 /*append*/);
    if (rc != MZ_OK) {
        m_lastError = QString("mz_zip_writer_open_file rc=%1").arg(rc);
        mz_zip_writer_delete(&m_handle);
        m_handle = nullptr;
        return;
    }
    m_open = true;
}

ZipStreamWriter::~ZipStreamWriter() {
    closeActiveEntry();
    if (m_open) {
        mz_zip_writer_close(m_handle);
        m_open = false;
    }
    if (m_handle) {
        mz_zip_writer_delete(&m_handle);
        m_handle = nullptr;
    }
}

void ZipStreamWriter::closeActiveEntry() {
    if (m_entryOpen) {
        mz_zip_writer_entry_close(m_handle);
        m_entryOpen = false;
    }
}

bool ZipStreamWriter::writeEntry(const QString& path, const QByteArray& data) {
    if (!m_open) return false;
    closeActiveEntry();

    mz_zip_file fi{};
    const QByteArray nameUtf8 = path.toUtf8();
    fi.filename = nameUtf8.constData();
    fi.compression_method = MZ_COMPRESS_METHOD_DEFLATE;
    fi.zip64 = MZ_ZIP64_AUTO;

    const int32_t rc = mz_zip_writer_add_buffer(m_handle,
                                                  const_cast<char*>(data.constData()),
                                                  data.size(),
                                                  &fi);
    if (rc != MZ_OK) {
        m_lastError = QString("mz_zip_writer_add_buffer rc=%1 for %2").arg(rc).arg(path);
        return false;
    }
    return true;
}

std::unique_ptr<ZipEntryWriteDevice> ZipStreamWriter::openEntry(const QString& path) {
    if (!m_open) return nullptr;
    closeActiveEntry();

    mz_zip_file fi{};
    const QByteArray nameUtf8 = path.toUtf8();
    fi.filename = nameUtf8.constData();
    fi.compression_method = MZ_COMPRESS_METHOD_DEFLATE;
    fi.zip64 = MZ_ZIP64_AUTO;

    const int32_t rc = mz_zip_writer_entry_open(m_handle, &fi);
    if (rc != MZ_OK) {
        m_lastError = QString("mz_zip_writer_entry_open rc=%1 for %2").arg(rc).arg(path);
        return nullptr;
    }
    m_entryOpen = true;

    std::unique_ptr<ZipEntryWriteDevice> dev(new ZipEntryWriteDevice(this));
    if (!dev->open(QIODevice::WriteOnly)) {
        m_lastError = "ZipEntryWriteDevice::open failed";
        closeActiveEntry();
        return nullptr;
    }
    return dev;
}

// ----- ZipEntryWriteDevice -----

ZipEntryWriteDevice::ZipEntryWriteDevice(ZipStreamWriter* owner)
    : m_owner(owner) {}

ZipEntryWriteDevice::~ZipEntryWriteDevice() {
    if (m_owner) {
        m_owner->closeActiveEntry();
    }
}

qint64 ZipEntryWriteDevice::writeData(const char* data, qint64 len) {
    if (!m_owner || !m_owner->m_entryOpen) return -1;
    if (len <= 0) return 0;

    // mz_zip_writer_entry_write takes int32_t length; loop for large writes.
    qint64 written = 0;
    while (written < len) {
        const qint64 remaining = len - written;
        const int32_t chunk = static_cast<int32_t>(qMin<qint64>(remaining, 1 << 30));
        const int32_t n = mz_zip_writer_entry_write(m_owner->m_handle,
                                                      data + written,
                                                      chunk);
        if (n < 0) return -1;
        if (n == 0) break;
        written += n;
    }
    return written;
}
