// Smoke test for ZipStreamReader.
// Creates an in-memory zip via minizip-ng's writer with known entries,
// writes it to a temp file, then reads it back through ZipStreamReader
// (both readEntry and the streaming QIODevice path) and asserts that
// the content round-trips byte-for-byte.

#include "ZipStreamReader.h"

#include "mz.h"
#include "mz_strm.h"
#include "mz_zip.h"
#include "mz_zip_rw.h"

#include <QCoreApplication>
#include <QFile>
#include <QTemporaryFile>
#include <QXmlStreamReader>
#include <QDebug>
#include <cstring>

namespace {

int g_failures = 0;

void check(bool cond, const char* what) {
    if (!cond) {
        qWarning() << "FAIL:" << what;
        ++g_failures;
    } else {
        qInfo() << "ok:  " << what;
    }
}

bool writeTestZip(const QString& outPath, const QList<QPair<QString, QByteArray>>& entries) {
    void* writer = mz_zip_writer_create();
    if (!writer) return false;

    int32_t rc = mz_zip_writer_open_file(writer, outPath.toUtf8().constData(), 0, 0);
    if (rc != MZ_OK) {
        mz_zip_writer_delete(&writer);
        return false;
    }

    for (const auto& kv : entries) {
        mz_zip_file fi{};
        const QByteArray nameUtf8 = kv.first.toUtf8();
        fi.filename = nameUtf8.constData();
        fi.compression_method = MZ_COMPRESS_METHOD_DEFLATE;
        fi.zip64 = MZ_ZIP64_AUTO;
        rc = mz_zip_writer_add_buffer(writer,
                                       const_cast<char*>(kv.second.constData()),
                                       kv.second.size(),
                                       &fi);
        if (rc != MZ_OK) {
            mz_zip_writer_close(writer);
            mz_zip_writer_delete(&writer);
            return false;
        }
    }
    mz_zip_writer_close(writer);
    mz_zip_writer_delete(&writer);
    return true;
}

} // namespace

int main(int argc, char** argv) {
    QCoreApplication app(argc, argv);

    QTemporaryFile tmp("nexel_zipstream_XXXXXX.zip");
    tmp.setAutoRemove(true);
    if (!tmp.open()) {
        qCritical() << "tmpfile open failed";
        return 1;
    }
    const QString zipPath = tmp.fileName();
    tmp.close();

    const QByteArray smallXml = "<?xml version=\"1.0\"?><root><a>1</a></root>";
    QByteArray largeXml;
    largeXml.append("<?xml version=\"1.0\"?><rows>");
    for (int i = 0; i < 20000; ++i) {
        largeXml.append(QString("<r i=\"%1\">value</r>").arg(i).toUtf8());
    }
    largeXml.append("</rows>");

    const bool wrote = writeTestZip(zipPath, {
        {"xl/workbook.xml",            smallXml},
        {"xl/worksheets/sheet1.xml",   largeXml},
        {"xl/sharedStrings.xml",       QByteArray("<sst/>")},
    });
    check(wrote, "wrote test zip");

    {
        ZipStreamReader r(zipPath);
        check(r.isOpen(), "reader opened");

        const QStringList ents = r.entries();
        check(ents.size() == 3, "entry count == 3");
        check(ents.contains("xl/workbook.xml"), "entries contain workbook.xml");
        check(ents.contains("xl/worksheets/sheet1.xml"), "entries contain sheet1.xml");

        const qint64 sz = r.entrySize("xl/worksheets/sheet1.xml");
        check(sz == largeXml.size(), "entrySize matches uncompressed size");

        const QByteArray got = r.readEntry("xl/workbook.xml");
        check(got == smallXml, "readEntry round-trips small file");

        const QByteArray gotLarge = r.readEntry("xl/worksheets/sheet1.xml");
        check(gotLarge == largeXml, "readEntry round-trips large file");
    }

    {
        ZipStreamReader r(zipPath);
        auto dev = r.openEntry("xl/worksheets/sheet1.xml");
        check(dev != nullptr, "openEntry returned device");
        check(dev && dev->isOpen(), "device is open");
        check(dev && dev->isSequential(), "device is sequential");

        QXmlStreamReader xml(dev.get());
        int rowCount = 0;
        while (!xml.atEnd()) {
            xml.readNext();
            if (xml.isStartElement() && xml.name() == QStringLiteral("r")) {
                ++rowCount;
            }
        }
        check(!xml.hasError(), "QXmlStreamReader parsed without error");
        check(rowCount == 20000, "all 20000 <r> elements seen via streaming");
    }

    {
        ZipStreamReader r(zipPath);
        auto dev = r.openEntry("xl/worksheets/sheet1.xml");
        check(dev != nullptr, "second openEntry on same reader");
        if (dev) {
            qint64 totalRead = 0;
            char buf[1024];
            while (!dev->atEnd()) {
                const qint64 n = dev->read(buf, sizeof(buf));
                if (n <= 0) break;
                totalRead += n;
            }
            check(totalRead == largeXml.size(), "chunked read totals uncompressed size");
        }
    }

    if (g_failures) {
        qCritical() << g_failures << "FAILURES";
        return 1;
    }
    qInfo() << "ALL PASS";
    return 0;
}
