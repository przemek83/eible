#ifndef IMPORTXLSXTEST_H
#define IMPORTXLSXTEST_H

#include <QObject>
#include <QSet>

class ImportXlsxTest : public QObject
{
    Q_OBJECT
private Q_SLOTS:
    void testRetrievingSheetNames();
    void testRetrievingSheetNamesFromEmptyFile();

    void testGetStyles();
    void testGetStylesNoContent();

    void testGetSharedStrings();
    void testGetSharedStringsNoContent();

private:
    static QMap<QString, QString> sheetNames_;

    static QStringList sharedStrings_;
};

#endif  // IMPORTXLSXTEST_H
