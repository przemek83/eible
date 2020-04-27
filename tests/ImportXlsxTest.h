#ifndef IMPORTXLSXTEST_H
#define IMPORTXLSXTEST_H

#include <QObject>
#include <QSet>

#include <ColumnType.h>

class ImportXlsx;

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

    void testGetData_data();
    void testGetData();

    void testGetDataLimitRows_data();
    void testGetDataLimitRows();

    void testGetDataExcludeColumns_data();
    void testGetDataExcludeColumns();
    void testGetDataExcludeInvalidColumn();

    void benchmarkGetData_data();
    void benchmarkGetData();
    void benchmarkGetData_clean();

private:
    QVector<QVector<QVariant>> getDataWithoutColumns(
        const QVector<QVector<QVariant>>& data, QVector<int> columnsToExclude);
    void setCommonData(ImportXlsx& importXlsx);
    static QList<std::pair<QString, QString>> sheets_;
    static QStringList sharedStrings_;
    static QList<int> dateStyles_;
    static QList<int> allStyles_;
    static QList<QStringList> testColumnNames_;
    static std::vector<unsigned int> expectedRowCounts_;
    static QVector<QVector<QVector<QVariant>>> sheetData_;
    static QVector<QVector<ColumnType>> columnTypes_;
    QByteArray benchmarkData_;
};

#endif  // IMPORTXLSXTEST_H
