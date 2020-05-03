#include "ImportOdsTest.h"

#include <ImportOds.h>
#include <QFile>
#include <QStringLiteral>
#include <QTest>

#include "ImportCommon.h"

const QString ImportOdsTest::testFileName_{QStringLiteral(":/testOds.ods")};

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
