#include "ImportOdsTest.h"

#include <ImportOds.h>
#include <QFile>
#include <QStringLiteral>

#include "ImportCommon.h"

void ImportOdsTest::testRetrievingSheetNames()
{
    ImportCommon().checkRetrievingSheetNames<ImportOds>(
        QStringLiteral(":/testOds.ods"));
}

void ImportOdsTest::testRetrievingSheetNamesFromEmptyFile() {}
