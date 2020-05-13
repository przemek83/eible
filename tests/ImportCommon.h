#ifndef IMPORTCOMMON_H
#define IMPORTCOMMON_H

#include <ImportSpreadsheet.h>
#include <QStringList>

class ImportCommon
{
public:
    template <class T>
    static void checkRetrievingSheetNames(const QString& fileName);

    template <class T>
    static void checkRetrievingSheetNamesFromEmptyFile();

    static void prepareDataForGetColumnListTest();
    template <class T>
    static void checkGetColumnList(const QString& fileName);

    template <class T>
    static void checkSettingEmptyColumnName(const QString& fileName);

    template <class T>
    static void checkGetColumnListTwoSheets(const QString& fileName);

    static void prepareDataForGetColumnTypes();
    template <class T>
    static void checkGetColumnTypes(const QString& fileName);

    static void prepareDataForGetColumnCountTest();
    template <class T>
    static void checkGetColumnCount(const QString& fileName);

    static void prepareDataForGetRowCountTest();
    template <class T>
    static void checkGetRowCount(const QString& fileName);

    static void prepareDataForGetRowAndColumnCountViaGetColumnTypes();
    template <class T>
    static void testGetRowAndColumnCountViaGetColumnTypes(
        const QString& fileName);

    static void prepareDataForGetData();
    template <class T>
    static void checkGetData(const QString& fileName);

    static void prepareDataForGetDataLimitRows();
    template <class T>
    static void checkGetDataLimitRows(const QString& fileName);

    static void prepareDataForGetGetDataExcludeColumns();
    template <class T>
    static void checkGetDataExcludeColumns(const QString& fileName);
    template <class T>
    static void checkGetDataExcludeInvalidColumn(const QString& fileName);

    static QVector<QVector<QVariant>> getDataForSheet(const QString& fileName);

    static QVector<QVector<QVariant>> getDataWithoutColumns(
        const QVector<QVector<QVariant>>& data, QVector<int> columnsToExclude);

    template <class T>
    static void checkEmittingProgressPercentChangedEmptyFile(
        const QString& fileName);
    template <class T>
    static void checkEmittingProgressPercentChangedSmallFile(
        const QString& fileName);
    template <class T>
    static void checkEmittingProgressPercentChangedBigFile(
        const QString& fileName);

    static QStringList getSheetNames();

    template <class T>
    static void checkInvalidSheetName(const QString& fileName);

private:
    static const QStringList sheetNames_;
    static const QList<QStringList> testColumnNames_;
    static const QVector<unsigned int> expectedRowCounts_;
    static const QVector<QVector<ColumnType>> columnTypes_;
    static const QVector<QVector<QVector<QVariant>>> sheetData_;

    static constexpr int NO_SIGNAL{0};
};

#endif  // IMPORTCOMMON_H
