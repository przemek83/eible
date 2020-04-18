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

    void testGetColumnList_data();
    void testGetColumnList();

private:
    static QMap<QString, QString> sheetNames_;

    static QStringList sharedStrings_;

    static QStringList testSheet1Columns_;
    static QStringList testSheet2Columns_;
    static QStringList testSheet3Columns_;
    static QStringList testSheet4Columns_;
    static QStringList testSheet5Columns_;
    static QStringList testSheet6Columns_;
    static QStringList testSheet7Columns_;
};

#endif  // IMPORTXLSXTEST_H
