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
};

#endif // EXPORTXLSXTEST
