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

const QList<QStringList> ImportCommon::testColumnNames_ = {
    {"Text", "Numeric", "Date"},
    {"Trait #1", "Value #1", "Transaction date", "Units", "Price",
     "Price per unit", "Another trait"},
    {},
    {"cena nier", "pow", "cena metra", "data transakcji", "text"},
    {"name", "date", "mass (kg)", "height", "---", "---", "---", "---", "---",
     "---", "---", "---"},
    {"modificator", "x", "y"},
    {"Pow", "Cena", "cena_m", "---", "---", "---"},
    {"user", "pass", "lic_exp", "uwagi"}};

const std::vector<unsigned int> ImportCommon::expectedRowCounts_{2,  19, 0,  4,
                                                                 30, 25, 20, 3};

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

void ImportCommon::prepareDataForGetColumnCountTest()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<int>("expectedColumnCount");
    for (int i = 0; i < testColumnNames_.size(); ++i)
    {
        const QString& sheetName{sheetNames_[i]};
        QTest::newRow(("Columns in " + sheetName).toStdString().c_str())
            << sheetName << testColumnNames_[i].size();
    }
}

template <class T>
void ImportCommon::checkGetColumnCount(const QString& fileName)
{
    QFETCH(QString, sheetName);
    QFETCH(int, expectedColumnCount);

    QFile testFile(fileName);
    T import(testFile);
    auto [success, actualColumnCount] = import.getColumnCount(sheetName);
    QCOMPARE(success, true);
    QCOMPARE(actualColumnCount, expectedColumnCount);
}
template void ImportCommon::checkGetColumnCount<ImportXlsx>(
    const QString& fileName);
template void ImportCommon::checkGetColumnCount<ImportOds>(
    const QString& fileName);

void ImportCommon::prepareDataForGetRowCountTest()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<unsigned int>("expectedRowCount");
    for (int i = 0; i < testColumnNames_.size(); ++i)
    {
        const QString& sheetName{sheetNames_[i]};
        QTest::newRow(("Rows in " + sheetName).toStdString().c_str())
            << sheetName << expectedRowCounts_[i];
    }
}

template <class T>
void ImportCommon::checkGetRowCount(const QString& fileName)
{
    QFETCH(QString, sheetName);
    QFETCH(unsigned int, expectedRowCount);

    QFile testFile(fileName);
    T import(testFile);
    auto [success, actualRowCount] = import.getRowCount(sheetName);
    QCOMPARE(success, true);
    QCOMPARE(actualRowCount, expectedRowCount);
}
template void ImportCommon::checkGetRowCount<ImportXlsx>(
    const QString& fileName);
template void ImportCommon::checkGetRowCount<ImportOds>(
    const QString& fileName);
