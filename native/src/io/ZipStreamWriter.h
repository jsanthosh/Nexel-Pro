#ifndef ZIPSTREAMWRITER_H
#define ZIPSTREAMWRITER_H

#include <QString>
#include <QIODevice>
#include <memory>

class ZipEntryWriteDevice;

class ZipStreamWriter {
public:
    explicit ZipStreamWriter(const QString& filePath);
    ~ZipStreamWriter();

    ZipStreamWriter(const ZipStreamWriter&) = delete;
    ZipStreamWriter& operator=(const ZipStreamWriter&) = delete;

    bool isOpen() const { return m_open; }
    QString lastError() const { return m_lastError; }

    // Small entries: write the whole buffer in one call. Returns false on error.
    bool writeEntry(const QString& path, const QByteArray& data);

    // Large entries: open a streaming sink. The returned device's writeData
    // pipes directly into the deflate stream; let it go out of scope to
    // finalize the entry before opening another one.
    std::unique_ptr<ZipEntryWriteDevice> openEntry(const QString& path);

private:
    friend class ZipEntryWriteDevice;

    void closeActiveEntry();

    void* m_handle = nullptr;
    bool m_open = false;
    bool m_entryOpen = false;
    QString m_lastError;
};

class ZipEntryWriteDevice : public QIODevice {
    Q_OBJECT
public:
    ~ZipEntryWriteDevice() override;

    bool isSequential() const override { return true; }

protected:
    qint64 readData(char*, qint64) override { return -1; }
    qint64 writeData(const char* data, qint64 len) override;

private:
    friend class ZipStreamWriter;
    explicit ZipEntryWriteDevice(ZipStreamWriter* owner);

    ZipStreamWriter* m_owner;
};

#endif
