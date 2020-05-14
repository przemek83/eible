#include "ImportOdsTest.h"

#include <ImportOds.h>
#include <QBuffer>
#include <QFile>
#include <QStringLiteral>
#include <QTest>

#include "ImportCommon.h"

ImportOdsTest::ImportOdsTest(QObject* parent) : QObject(parent) {}

void ImportOdsTest::testRetrievingSheetNames()
{
    QFile testFile(testFileName_);
    ImportOds importOds(testFile);
    ImportCommon::checkRetrievingSheetNames(importOds);
}

void ImportOdsTest::testRetrievingSheetNamesFromEmptyFile()
{
    QByteArray byteArray;
    QBuffer emptyBuffer(&byteArray);
    ImportOds importOds(emptyBuffer);
    ImportCommon::checkRetrievingSheetNamesFromEmptyFile(importOds);
}

void ImportOdsTest::testGetColumnList_data()
{
    ImportCommon::prepareDataForGetColumnListTest();
}

void ImportOdsTest::testGetColumnList()
{
    QFile testFile(testFileName_);
    ImportOds importOds(testFile);
    ImportCommon::checkGetColumnList(importOds);
}

void ImportOdsTest::testSettingEmptyColumnName()
{
    QFile testFile(testFileName_);
    ImportOds importOds(testFile);
    ImportCommon::checkSettingEmptyColumnName(importOds);
}

void ImportOdsTest::testGetColumnListTwoSheets()
{
    QFile testFile(testFileName_);
    ImportOds importOds(testFile);
    ImportCommon::checkGetColumnListTwoSheets(importOds);
}

void ImportOdsTest::testGetColumnTypes_data()
{
    ImportCommon::prepareDataForGetColumnTypes();
}

void ImportOdsTest::testGetColumnTypes()
{
    QFile testFile(testFileName_);
    ImportOds importOds(testFile);
    ImportCommon::checkGetColumnTypes(importOds);
}

void ImportOdsTest::testGetColumnCount_data()
{
    ImportCommon::prepareDataForGetColumnCountTest();
}

void ImportOdsTest::testGetColumnCount()
{
    QFile testFile(testFileName_);
    ImportOds importOds(testFile);
    ImportCommon::checkGetColumnCount(importOds);
}

void ImportOdsTest::testGetRowCount_data()
{
    ImportCommon::prepareDataForGetRowCountTest();
}

void ImportOdsTest::testGetRowCount()
{
    QFile testFile(testFileName_);
    ImportOds importOds(testFile);
    ImportCommon::checkGetRowCount(importOds);
}

void ImportOdsTest::testGetRowAndColumnCountViaGetColumnTypes_data()
{
    ImportCommon::prepareDataForGetRowAndColumnCountViaGetColumnTypes();
}

void ImportOdsTest::testGetRowAndColumnCountViaGetColumnTypes()
{
    QFile testFile(testFileName_);
    ImportOds importOds(testFile);
    ImportCommon::testGetRowAndColumnCountViaGetColumnTypes(importOds);
}

void ImportOdsTest::testGetData_data()
{
    ImportCommon::prepareDataForGetData();
}

void ImportOdsTest::testGetData()
{
    QFile testFile(testFileName_);
    ImportOds importOds(testFile);
    ImportCommon::checkGetData(importOds);
}

void ImportOdsTest::testGetDataLimitRows_data()
{
    ImportCommon::prepareDataForGetDataLimitRows();
}

void ImportOdsTest::testGetDataLimitRows()
{
    QFile testFile(testFileName_);
    ImportOds importOds(testFile);
    ImportCommon::checkGetDataLimitRows(importOds);
}

void ImportOdsTest::testGetDataExcludeColumns_data()
{
    ImportCommon::prepareDataForGetGetDataExcludeColumns();
}

void ImportOdsTest::testGetDataExcludeColumns()
{
    QFile testFile(testFileName_);
    ImportOds importOds(testFile);
    ImportCommon::checkGetDataExcludeColumns(importOds);
}

void ImportOdsTest::testGetDataExcludeInvalidColumn()
{
    QFile testFile(testFileName_);
    ImportOds importOds(testFile);
    ImportCommon::checkGetDataExcludeInvalidColumn(importOds);
}

void ImportOdsTest::benchmarkGetData()
{
    QSKIP("Skip benchmark.");
    QFile testFile(QStringLiteral(":/bigFile.ods"));
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
    QFile testFile(templateFileName_);
    ImportOds importOds(testFile);
    ImportCommon::checkEmittingProgressPercentChangedEmptyFile(importOds);
}

void ImportOdsTest::testEmittingProgressPercentChangedSmallFile()
{
    QFile testFile(testFileName_);
    ImportOds importOds(testFile);
    ImportCommon::checkEmittingProgressPercentChangedSmallFile(importOds);
}

void ImportOdsTest::testEmittingProgressPercentChangedBigFile()
{
    QFile testFile(QStringLiteral(":/mediumFile.ods"));
    ImportOds importOds(testFile);
    ImportCommon::checkEmittingProgressPercentChangedBigFile(importOds);
}

void ImportOdsTest::testInvalidSheetName()
{
    QFile testFile(testFileName_);
    ImportOds importOds(testFile);
    ImportCommon::checkInvalidSheetName(importOds);
}
