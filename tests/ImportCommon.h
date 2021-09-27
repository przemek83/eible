#pragma once

#include <ImportSpreadsheet.h>
#include <QStringList>
#include <QVariant>

namespace ImportCommon
{
void checkRetrievingSheetNames(ImportSpreadsheet& importer);

void checkRetrievingSheetNamesFromEmptyFile(ImportSpreadsheet& importer);

void prepareDataForGetColumnListTest();

void checkGetColumnList(ImportSpreadsheet& importer);

void checkSettingEmptyColumnName(ImportSpreadsheet& importer);

void checkGetColumnListTwoSheets(ImportSpreadsheet& importer);

void prepareDataForGetColumnTypes();

void checkGetColumnTypes(ImportSpreadsheet& importer);

void prepareDataForGetColumnCountTest();

void checkGetColumnCount(ImportSpreadsheet& importer);

void prepareDataForGetRowCountTest();

void checkGetRowCount(ImportSpreadsheet& importer);

void prepareDataForGetRowAndColumnCountViaGetColumnTypes();

void testGetRowAndColumnCountViaGetColumnTypes(ImportSpreadsheet& importer);

void prepareDataForGetData();

void checkGetData(ImportSpreadsheet& importer);

void prepareDataForGetDataLimitRows();

void checkGetDataLimitRows(ImportSpreadsheet& importer);

void prepareDataForGetGetDataExcludeColumns();

void checkGetDataExcludeColumns(ImportSpreadsheet& importer);

void checkGetDataExcludeInvalidColumn(ImportSpreadsheet& importer);

QVector<QVector<QVariant>> getDataForSheet(const QString& fileName);

QVector<QVector<QVariant>> getDataWithoutColumns(
    const QVector<QVector<QVariant>>& data,
    const QVector<int>& columnsToExclude);

void checkEmittingProgressPercentChangedEmptyFile(ImportSpreadsheet& importer);

void checkEmittingProgressPercentChangedSmallFile(ImportSpreadsheet& importer);

void checkEmittingProgressPercentChangedBigFile(ImportSpreadsheet& importer);

QStringList& getSheetNames();

void checkInvalidSheetName(ImportSpreadsheet& importer);
};  // namespace ImportCommon
