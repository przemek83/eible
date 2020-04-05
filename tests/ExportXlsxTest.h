#ifndef EXPORTXLSXTEST
#define EXPORTXLSXTEST

#include <QObject>
#include <QTableWidget>

class QTableWidget;

class ExportXlsxTest: public QObject
{
    Q_OBJECT
private Q_SLOTS:
    void initTestCase();

    void testExportingEmptyTable();

    void testExportingHeadersOnly();

    void testExportingSimpleTable();

    void benchmark();

private:
    QByteArray retrieveFileFromZip(const QString& zipFilePath,
                                   const QString& fileName) const;

    void initTable(QTableWidget& tableWidget, int columnCount, int rowCount) const;

    void compareWorkSheets(const QString& testFilePath,
                           const QString& sheetData) const;

    void exportZip(const QTableWidget& tableWidget,
                   const QString& testFile) const;

    static QString zipWorkSheetPath_;
    static QString tableSheetData_;
    static QString headersOnlySheetData_;
    static QString emptySheetData_;
    static QStringList headers_;
    QTableWidget tableWidgetForBenchmarking_;
};

#endif // EXPORTXLSXTEST
