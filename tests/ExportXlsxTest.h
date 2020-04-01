#ifndef EXPORTXLSXTEST
#define EXPORTXLSXTEST

#include <QObject>

class QTableWidget;

class ExportXlsxTest: public QObject
{
    Q_OBJECT
private Q_SLOTS:
    void testExportingEmptyTable();

private:
    QByteArray retrieveFileFromZip(const QString& zipFilePath,
                                   const QString& fileName) const;

    QByteArray computeHash(const QByteArray& fileContent) const;

    void initTable(QTableWidget& tableWidget) const;
};

#endif // EXPORTXLSXTEST
