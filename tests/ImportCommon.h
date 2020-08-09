#pragma once

#include <ImportSpreadsheet.h>
#include <QStringList>

class ImportCommon
{
public:
    static void checkRetrievingSheetNames(ImportSpreadsheet& importer);

    static void checkRetrievingSheetNamesFromEmptyFile(
        ImportSpreadsheet& importer);

    static void prepareDataForGetColumnListTest();
    static void checkGetColumnList(ImportSpreadsheet& importer);

    static void checkSettingEmptyColumnName(ImportSpreadsheet& importer);

    static void checkGetColumnListTwoSheets(ImportSpreadsheet& importer);

    static void prepareDataForGetColumnTypes();
    static void checkGetColumnTypes(ImportSpreadsheet& importer);

    static void prepareDataForGetColumnCountTest();
    static void checkGetColumnCount(ImportSpreadsheet& importer);

    static void prepareDataForGetRowCountTest();
    static void checkGetRowCount(ImportSpreadsheet& importer);

    static void prepareDataForGetRowAndColumnCountViaGetColumnTypes();
    static void testGetRowAndColumnCountViaGetColumnTypes(
        ImportSpreadsheet& importer);

    static void prepareDataForGetData();
    static void checkGetData(ImportSpreadsheet& importer);

    static void prepareDataForGetDataLimitRows();
    static void checkGetDataLimitRows(ImportSpreadsheet& importer);

    static void prepareDataForGetGetDataExcludeColumns();
    static void checkGetDataExcludeColumns(ImportSpreadsheet& importer);
    static void checkGetDataExcludeInvalidColumn(ImportSpreadsheet& importer);

    static QVector<QVector<QVariant>> getDataForSheet(const QString& fileName);

    static QVector<QVector<QVariant>> getDataWithoutColumns(
        const QVector<QVector<QVariant>>& data, QVector<int> columnsToExclude);

    static void checkEmittingProgressPercentChangedEmptyFile(
        ImportSpreadsheet& importer);
    static void checkEmittingProgressPercentChangedSmallFile(
        ImportSpreadsheet& importer);
    static void checkEmittingProgressPercentChangedBigFile(
        ImportSpreadsheet& importer);

    static QStringList getSheetNames();

    static void checkInvalidSheetName(ImportSpreadsheet& importer);

private:
    static const QStringList sheetNames_;
    static const QList<QStringList> testColumnNames_;
    static const QVector<unsigned int> expectedRowCounts_;
    static const QVector<QVector<ColumnType>> columnTypes_;
    static const QVector<QVector<QVector<QVariant>>> sheetData_;

    static constexpr int NO_SIGNAL{0};
};
