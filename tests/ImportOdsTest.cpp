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

void ImportOdsTest::testSettingEmptyColumnName() {}

void ImportOdsTest::testGetColumnListTwoSheets() {}

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
