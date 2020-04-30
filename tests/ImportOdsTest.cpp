#include "ImportOdsTest.h"

#include <ImportOds.h>
#include <QFile>
#include <QStringLiteral>

#include "ImportCommon.h"

const QString ImportOdsTest::testFileName_{QStringLiteral(":/testOds.ods")};

void ImportOdsTest::testRetrievingSheetNames()
{
    ImportCommon().checkRetrievingSheetNames<ImportOds>(testFileName_);
}

void ImportOdsTest::testRetrievingSheetNamesFromEmptyFile()
{
    ImportCommon().checkRetrievingSheetNamesFromEmptyFile<ImportOds>();
}
