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
    QSKIP("TODO");
    ImportCommon::prepareDataForGetRowCountTest();
}

void ImportOdsTest::testGetRowCount()
{
    QSKIP("TODO");
    ImportCommon::checkGetRowCount<ImportOds>(testFileName_);
}
