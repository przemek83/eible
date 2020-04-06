#ifndef EXPORTXLSXTEST
#define EXPORTXLSXTEST

#include <QObject>

class QAbstractItemView;
class TestTableModel;
class QBuffer;

class ExportXlsxTest: public QObject
{
    Q_OBJECT
private Q_SLOTS:
    void initTestCase();

    void testExportingEmptyTable();

    void testExportingHeadersOnly();

    void testExportingSimpleTable();

    void testExportingViewWithMultiSelection();

    void Benchmark_data();

    void Benchmark();

    void cleanupTestCase();

private:
    QByteArray retrieveFileFromZip(QBuffer& exportedZip,
                                   const QString& fileName) const;

    void compareWorkSheets(QBuffer& exportedZip,
                           const QString& sheetData) const;

    void exportZip(const QAbstractItemView* view,
                   QBuffer& exportedZip) const;

    static QString zipWorkSheetPath_;
    static QString tableSheetData_;
    static QString multiSelectionTableSheetData_;
    static QString headersOnlySheetData_;
    static QString emptySheetData_;
    static QStringList headers_;
    TestTableModel* tableModelForBenchmarking_ {nullptr};
};

#endif // EXPORTXLSXTEST
