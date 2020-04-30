#include "ImportCommon.h"

#include <ImportOds.h>
#include <ImportXlsx.h>
#include <QBuffer>
#include <QFile>
#include <QTest>

#include "ImportSpreadsheet.h"

const QStringList ImportCommon::sheetNames_{
    "Sheet1", "Sheet2", "Sheet3(empty)", "Sheet4",
    "Sheet5", "Sheet6", "Sheet7",        "testAccounts"};

template <class T>
void ImportCommon::checkRetrievingSheetNames(const QString& fileName)
{
    QFile testFile(fileName);
    T import(testFile);
    auto [success, actualSheetNames] = import.getSheetNames();
    QCOMPARE(success, true);
    QCOMPARE(actualSheetNames, sheetNames_);
}
template void ImportCommon::checkRetrievingSheetNames<ImportXlsx>(
    const QString& fileName);
template void ImportCommon::checkRetrievingSheetNames<ImportOds>(
    const QString& fileName);

template <class T>
void ImportCommon::checkRetrievingSheetNamesFromEmptyFile()
{
    QByteArray byteArray;
    QBuffer emptyBuffer(&byteArray);
    T import(emptyBuffer);
    auto [success, actualNames] = import.getSheetNames();
    QCOMPARE(success, false);
    QCOMPARE(actualNames, {});
}
template void
ImportCommon::checkRetrievingSheetNamesFromEmptyFile<ImportXlsx>();
template void ImportCommon::checkRetrievingSheetNamesFromEmptyFile<ImportOds>();
