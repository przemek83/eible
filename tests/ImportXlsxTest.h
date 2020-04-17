#ifndef IMPORTXLSXTEST_H
#define IMPORTXLSXTEST_H

#include <QObject>

class ImportXlsxTest : public QObject
{
    Q_OBJECT
private Q_SLOTS:
    void testRetrievingSheetNames();
    void testRetrievingSheetNamesFromEmptyFile();

    void testGetStyles();
    void testGetStylesNoContent();

private:
    static QMap<QString, QString> sheetNames_;
};

#endif  // IMPORTXLSXTEST_H
