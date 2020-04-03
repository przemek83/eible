#ifndef EXPORTXLSXTEST
#define EXPORTXLSXTEST

#include <QObject>

class QTableWidget;

class ExportXlsxTest: public QObject
{
    Q_OBJECT
private Q_SLOTS:
    void testExportingEmptyTable();

    void testExportingHeadersOnly();

    void testExportingSimpleTable();

private:
    QByteArray retrieveFileFromZip(const QString& zipFilePath,
                                   const QString& fileName) const;

    void initTable(QTableWidget& tableWidget) const;

    void compareWorkSheets(const QString& testFilePath,
                           const QString& sheetData) const;

    void exportZip(const QTableWidget& tableWidget,
                   const QString& testFile) const;

    static QString zipWorkSheetPath_;
    static QString tableSheetData_;
    static QString headersOnlySheetData_;
    static QString emptySheetData_;
    static QStringList headers_;
};

#endif // EXPORTXLSXTEST
