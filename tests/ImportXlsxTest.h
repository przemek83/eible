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

    void testGetDateStyles();
    void testGetDateStylesNoContent();

    void testGetAllStyles();
    void testGetAllStylesNoContent();

    void testGetSharedStrings();
    void testGetSharedStringsNoContent();

    void testGetColumnList_data();
    void testGetColumnList();
    void testSettingEmptyColumnName();
    void testGetColumnListTwoSheets();

    void testGetColumnTypes_data();
    void testGetColumnTypes();

    void testGetColumnCount_data();
    void testGetColumnCount();

    void testGetRowCount_data();
    void testGetRowCount();

    void testGetRowAndColumnCountViaGetColumnTypes_data();
    void testGetRowAndColumnCountViaGetColumnTypes();

private:
    static QList<std::pair<QString, QString>> sheets_;
    static QStringList sharedStrings_;
    static QList<int> dateStyles_;
    static QList<int> allStyles_;
    static QList<QStringList> testColumnNames_;
    static std::vector<unsigned int> expectedRowCounts_;
};

#endif  // IMPORTXLSXTEST_H
