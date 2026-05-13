#ifndef ZIPSTREAMREADER_H
#define ZIPSTREAMREADER_H

#include <QString>
#include <QStringList>
#include <QIODevice>
#include <memory>

class ZipEntryDevice;

class ZipStreamReader {
public:
    explicit ZipStreamReader(const QString& filePath);
    ~ZipStreamReader();

    ZipStreamReader(const ZipStreamReader&) = delete;
    ZipStreamReader& operator=(const ZipStreamReader&) = delete;

    bool isOpen() const { return m_open; }
    QString lastError() const { return m_lastError; }

    QStringList entries() const;
    qint64 entrySize(const QString& path) const;

    std::unique_ptr<ZipEntryDevice> openEntry(const QString& path);
    QByteArray readEntry(const QString& path);

private:
    friend class ZipEntryDevice;

    void closeActiveEntry();

    void* m_handle = nullptr;
    bool m_open = false;
    bool m_entryOpen = false;
    QString m_lastError;
};

class ZipEntryDevice : public QIODevice {
    Q_OBJECT
public:
    ~ZipEntryDevice() override;

    bool isSequential() const override { return true; }
    qint64 size() const override { return m_uncompressedSize; }
    qint64 bytesAvailable() const override;

protected:
    qint64 readData(char* data, qint64 maxlen) override;
    qint64 writeData(const char*, qint64) override { return -1; }

private:
    friend class ZipStreamReader;
    ZipEntryDevice(ZipStreamReader* owner, qint64 uncompressedSize);

    ZipStreamReader* m_owner;
    qint64 m_uncompressedSize;
    qint64 m_bytesRead = 0;
};

#endif
