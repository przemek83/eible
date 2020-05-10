#include "ImportOdsTest.h"

#include <ImportOds.h>
#include <QFile>
#include <QStringLiteral>
#include <QTest>

#include "ImportCommon.h"

const QString ImportOdsTest::testFileName_{QStringLiteral(":/testOds.ods")};
const QString ImportOdsTest::templateFileName_{
    QStringLiteral(":/emptyOds.ods")};

void ImportOdsTest::testRetrievingSheetNames()
{
    ImportCommon::checkRetrievingSheetNames<ImportOds>(testFileName_);
}

void ImportOdsTest::testRetrievingSheetNamesFromEmptyFile()
{
    ImportCommon::checkRetrievingSheetNamesFromEmptyFile<ImportOds>();
}

void ImportOdsTest::testGetColumnList_data()
{
    ImportCommon::prepareDataForGetColumnListTest();
}

void ImportOdsTest::testGetColumnList()
{
    ImportCommon::checkGetColumnList<ImportOds>(testFileName_);
}

void ImportOdsTest::testSettingEmptyColumnName()
{
    ImportCommon::checkSettingEmptyColumnName<ImportOds>(testFileName_);
}

void ImportOdsTest::testGetColumnListTwoSheets()
{
    ImportCommon::checkGetColumnListTwoSheets<ImportOds>(testFileName_);
}

void ImportOdsTest::testGetColumnTypes_data()
{
    ImportCommon::prepareDataForGetColumnTypes();
}

void ImportOdsTest::testGetColumnTypes()
{
    ImportCommon::checkGetColumnTypes<ImportOds>(testFileName_);
}

void ImportOdsTest::testGetColumnCount_data()
{
    ImportCommon::prepareDataForGetColumnCountTest();
}

void ImportOdsTest::testGetColumnCount()
{
    ImportCommon::checkGetColumnCount<ImportOds>(testFileName_);
}

void ImportOdsTest::testGetRowCount_data()
{
    ImportCommon::prepareDataForGetRowCountTest();
}

void ImportOdsTest::testGetRowCount()
{
    ImportCommon::checkGetRowCount<ImportOds>(testFileName_);
}

void ImportOdsTest::testGetRowAndColumnCountViaGetColumnTypes_data()
{
    ImportCommon::prepareDataForGetRowAndColumnCountViaGetColumnTypes();
}

void ImportOdsTest::testGetRowAndColumnCountViaGetColumnTypes()
{
    ImportCommon::testGetRowAndColumnCountViaGetColumnTypes<ImportOds>(
        testFileName_);
}

void ImportOdsTest::testGetData_data()
{
    ImportCommon::prepareDataForGetData();
}

void ImportOdsTest::testGetData()
{
    ImportCommon::checkGetData<ImportOds>(testFileName_);
}

void ImportOdsTest::testGetDataLimitRows_data()
{
    ImportCommon::prepareDataForGetDataLimitRows();
}

void ImportOdsTest::testGetDataLimitRows()
{
    ImportCommon::checkGetDataLimitRows<ImportOds>(testFileName_);
}

void ImportOdsTest::testGetDataExcludeColumns_data()
{
    ImportCommon::prepareDataForGetGetDataExcludeColumns();
}

void ImportOdsTest::testGetDataExcludeColumns()
{
    ImportCommon::checkGetDataExcludeColumns<ImportOds>(testFileName_);
}

void ImportOdsTest::testGetDataExcludeInvalidColumn()
{
    ImportCommon::checkGetDataExcludeInvalidColumn<ImportOds>(testFileName_);
}

void ImportOdsTest::benchmarkGetData()
{
    QSKIP("Skip benchmark.");
    QFile testFile(":/bigFile.ods");
    ImportOds importOds(testFile);
    auto [success, sheetNames] = importOds.getSheetNames();
    QCOMPARE(success, true);
    QCOMPARE(sheetNames.size(), 1);
    QBENCHMARK
    {
        for (int i = 0; i < 10; ++i)
            importOds.getData(sheetNames.front(), {});
    }
}

void ImportOdsTest::testEmittingProgressPercentChangedEmptyFile()
{
    ImportCommon::checkEmittingProgressPercentChangedEmptyFile<ImportOds>(
        templateFileName_);
}

void ImportOdsTest::testEmittingProgressPercentChangedSmallFile()
{
    ImportCommon::checkEmittingProgressPercentChangedSmallFile<ImportOds>(
        testFileName_);
}

void ImportOdsTest::testEmittingProgressPercentChangedBigFile()
{
    ImportCommon::checkEmittingProgressPercentChangedBigFile<ImportOds>(
        ":/mediumFile.ods");
}

void ImportOdsTest::testInvalidSheetName()
{
    ImportCommon::checkInvalidSheetName<ImportOds>(testFileName_);
}
