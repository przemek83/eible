#include "ImportXlsxTest.h"

#include <ExportXlsx.h>
#include <ImportXlsx.h>
#include <QBuffer>
#include <QTableView>
#include <QTest>

#include "ImportCommon.h"
#include "TestTableModel.h"

ImportXlsxTest::ImportXlsxTest(QObject* parent) : QObject(parent) {}

void ImportXlsxTest::testRetrievingSheetNames()
{
    QFile testFile(testFileName_);
    ImportXlsx importXlsx(testFile);
    ImportCommon::checkRetrievingSheetNames(importXlsx);
}

void ImportXlsxTest::testRetrievingSheetNamesFromEmptyFile()
{
    QByteArray byteArray;
    QBuffer emptyBuffer(&byteArray);
    ImportXlsx importXlsx(emptyBuffer);
    ImportCommon::checkRetrievingSheetNamesFromEmptyFile(importXlsx);
}

void ImportXlsxTest::testGetDateStyles()
{
    QFile xlsxTestFile(testFileName_);
    ImportXlsx importXlsx(xlsxTestFile);
    auto [success, actualDateStyle] = importXlsx.getDateStyles();
    QCOMPARE(success, true);
    QCOMPARE(actualDateStyle, dateStyles_);
}

void ImportXlsxTest::testGetDateStylesNoContent()
{
    QFile xlsxTestFile(templateFileName_);
    ImportXlsx importXlsx(xlsxTestFile);
    auto [success, actualDateStyle] = importXlsx.getDateStyles();
    QCOMPARE(success, true);
    QCOMPARE(actualDateStyle, QList({14, 15, 16, 17, 22}));
}

void ImportXlsxTest::testGetAllStyles()
{
    QFile xlsxTestFile(testFileName_);
    ImportXlsx importXlsx(xlsxTestFile);
    auto [success, actualAllStyle] = importXlsx.getAllStyles();
    QCOMPARE(success, true);
    QCOMPARE(actualAllStyle, allStyles_);
}

void ImportXlsxTest::testGetAllStylesNoContent()
{
    QFile xlsxTestFile(templateFileName_);
    ImportXlsx importXlsx(xlsxTestFile);
    auto [success, actualAllStyle] = importXlsx.getAllStyles();
    QCOMPARE(success, true);
    QCOMPARE(actualAllStyle, QList({0, 164, 10, 14, 4, 0, 0, 3}));
}

void ImportXlsxTest::testGetSharedStrings()
{
    QFile xlsxTestFile(testFileName_);
    ImportXlsx importXlsx(xlsxTestFile);
    auto [success, actualSharedStrings] = importXlsx.getSharedStrings();
    QCOMPARE(success, true);
    QCOMPARE(actualSharedStrings, sharedStrings_);
}

void ImportXlsxTest::testGetSharedStringsNoContent()
{
    QFile xlsxTestFile(templateFileName_);
    ImportXlsx importXlsx(xlsxTestFile);
    auto [success, actualSharedStrings] = importXlsx.getSharedStrings();
    QCOMPARE(success, true);
    QCOMPARE(actualSharedStrings, {});
}

void ImportXlsxTest::testGetColumnList_data()
{
    ImportCommon::prepareDataForGetColumnListTest();
}

void ImportXlsxTest::testGetColumnList()
{
    QFile testFile(testFileName_);
    ImportXlsx importXlsx(testFile);
    ImportCommon::checkGetColumnList(importXlsx);
}

void ImportXlsxTest::testSettingEmptyColumnName()
{
    QFile testFile(testFileName_);
    ImportXlsx importXlsx(testFile);
    ImportCommon::checkSettingEmptyColumnName(importXlsx);
}

void ImportXlsxTest::testGetColumnListTwoSheets()
{
    QFile testFile(testFileName_);
    ImportXlsx importXlsx(testFile);
    ImportCommon::checkGetColumnListTwoSheets(importXlsx);
}

void ImportXlsxTest::testGetColumnTypes_data()
{
    ImportCommon::prepareDataForGetColumnTypes();
}

void ImportXlsxTest::testGetColumnTypes()
{
    QFile testFile(testFileName_);
    ImportXlsx importXlsx(testFile);
    ImportCommon::checkGetColumnTypes(importXlsx);
}

void ImportXlsxTest::testGetColumnCount_data()
{
    ImportCommon::prepareDataForGetColumnCountTest();
}

void ImportXlsxTest::testGetColumnCount()
{
    QFile testFile(testFileName_);
    ImportXlsx importXlsx(testFile);
    ImportCommon::checkGetColumnCount(importXlsx);
}

void ImportXlsxTest::testGetRowCount_data()
{
    ImportCommon::prepareDataForGetRowCountTest();
}

void ImportXlsxTest::testGetRowCount()
{
    QFile testFile(testFileName_);
    ImportXlsx importXlsx(testFile);
    ImportCommon::checkGetRowCount(importXlsx);
}

void ImportXlsxTest::testGetRowAndColumnCountViaGetColumnTypes_data()
{
    ImportCommon::prepareDataForGetRowAndColumnCountViaGetColumnTypes();
}

void ImportXlsxTest::testGetRowAndColumnCountViaGetColumnTypes()
{
    QFile testFile(testFileName_);
    ImportXlsx importXlsx(testFile);
    ImportCommon::testGetRowAndColumnCountViaGetColumnTypes(importXlsx);
}

QVector<QVector<QVariant>> ImportXlsxTest::convertDataToUseSharedStrings(
    const QVector<QVector<QVariant>>& inputData)
{
    QVector<QVector<QVariant>> outputData{inputData};
    for (auto& row : outputData)
        for (auto& item : row)
            if (item.type() == QVariant::String && !item.isNull())
                item = (sharedStrings_.contains(item.toString())
                            ? QVariant(sharedStrings_.indexOf(item.toString()))
                            : item);
    return outputData;
}

void ImportXlsxTest::testGetData_data()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<QVector<QVector<QVariant>>>("expectedData");
    QCOMPARE(sheets_.size(), ImportCommon::getSheetNames().size());
    for (const auto& [sheetName, sheetPath] : sheets_)
    {
        QVector<QVector<QVariant>> sheetData{
            ImportCommon::getDataForSheet(sheetName)};
        sheetData = convertDataToUseSharedStrings(sheetData);
        QTest::newRow(("Data in " + sheetName).toStdString().c_str())
            << sheetName << sheetData;
    }
}

void ImportXlsxTest::testGetData()
{
    QFile testFile(testFileName_);
    ImportXlsx importXlsx(testFile);
    ImportCommon::checkGetData(importXlsx);
}

void ImportXlsxTest::testGetDataLimitRows_data()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<unsigned int>("rowLimit");
    QTest::addColumn<QVector<QVector<QVariant>>>("expectedData");

    QString sheetName{sheets_[0].first};
    QVector<QVector<QVariant>> sheetData{convertDataToUseSharedStrings(
        ImportCommon::getDataForSheet(sheetName))};
    unsigned int rowLimit{10};
    QString testName{"Limited data to " + QString::number(rowLimit) + " in " +
                     sheetName};
    QTest::newRow(testName.toStdString().c_str())
        << sheetName << rowLimit << sheetData;

    rowLimit = 2U;
    testName =
        "Limited data to " + QString::number(rowLimit) + " in " + sheetName;
    QTest::newRow(testName.toStdString().c_str())
        << sheetName << rowLimit << sheetData;

    sheetName = sheets_[5].first;
    rowLimit = 12U;
    QVector<QVector<QVariant>> expectedValues;
    expectedValues.reserve(rowLimit);
    sheetData = ImportCommon::getDataForSheet(sheetName);
    for (unsigned int i = 0; i < rowLimit; ++i)
        expectedValues.append(sheetData[static_cast<int>(i)]);
    expectedValues = convertDataToUseSharedStrings(expectedValues);
    testName =
        "Limited data to " + QString::number(rowLimit) + " in " + sheetName;
    QTest::newRow(testName.toStdString().c_str())
        << sheetName << rowLimit << expectedValues;
}

void ImportXlsxTest::testGetDataLimitRows()
{
    QFile testFile(testFileName_);
    ImportXlsx importXlsx(testFile);
    ImportCommon::checkGetDataLimitRows(importXlsx);
}

void ImportXlsxTest::testGetDataExcludeColumns_data()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<QVector<unsigned int>>("excludedColumns");
    QTest::addColumn<QVector<QVector<QVariant>>>("expectedData");

    QString sheetName{sheets_[0].first};
    QString testName{"Get data with excluded column 1 in " + sheetName};
    QVector<QVector<QVariant>> sheetData{convertDataToUseSharedStrings(
        ImportCommon::getDataForSheet(sheetName))};
    QTest::newRow(qUtf8Printable(testName))
        << sheetName << QVector<unsigned int>{1}
        << ImportCommon::getDataWithoutColumns(sheetData, {1});

    testName = "Get data with excluded column 0 and 2 in " + sheetName;
    QTest::newRow(qUtf8Printable(testName))
        << sheetName << QVector<unsigned int>{0, 2}
        << ImportCommon::getDataWithoutColumns(sheetData, {2, 0});
}

void ImportXlsxTest::testGetDataExcludeColumns()
{
    QFile testFile(testFileName_);
    ImportXlsx importXlsx(testFile);
    ImportCommon::checkGetDataExcludeColumns(importXlsx);
}

void ImportXlsxTest::testGetDataExcludeInvalidColumn()
{
    QFile testFile(testFileName_);
    ImportXlsx importXlsx(testFile);
    ImportCommon::checkGetDataExcludeInvalidColumn(importXlsx);
}

void ImportXlsxTest::benchmarkGetData_data()
{
    QSKIP("Skip preparing benchmark data.");
    TestTableModel model(100, 100000);
    QTableView view;
    view.setModel(&model);

    generatedXlsx_.clear();
    QBuffer exportedFile(&generatedXlsx_);
    exportedFile.open(QIODevice::WriteOnly);
    ExportXlsx exportXlsx;
    exportXlsx.exportView(view, exportedFile);
}

void ImportXlsxTest::benchmarkGetData()
{
    QSKIP("Skip benchmark.");
    QBuffer testFile(&generatedXlsx_);
    ImportXlsx importXlsx(testFile);
    auto [success, sheetNames] = importXlsx.getSheetNames();
    QCOMPARE(success, true);
    QCOMPARE(sheetNames.size(), 1);
    QBENCHMARK { importXlsx.getData(sheetNames.front(), {}); }
}

void ImportXlsxTest::testEmittingProgressPercentChangedEmptyFile()
{
    QFile testFile(templateFileName_);
    ImportXlsx importXlsx(testFile);
    ImportCommon::checkEmittingProgressPercentChangedEmptyFile(importXlsx);
}

void ImportXlsxTest::testEmittingProgressPercentChangedSmallFile()
{
    QFile testFile(testFileName_);
    ImportXlsx importXlsx(testFile);
    ImportCommon::checkEmittingProgressPercentChangedSmallFile(importXlsx);
}

void ImportXlsxTest::testEmittingProgressPercentChangedBigFile()
{
    QFile testFile(QStringLiteral(":/mediumFile.xlsx"));
    ImportXlsx importXlsx(testFile);
    ImportCommon::checkEmittingProgressPercentChangedBigFile(importXlsx);
}

void ImportXlsxTest::testInvalidSheetName()
{
    QFile testFile(testFileName_);
    ImportXlsx importXlsx(testFile);
    ImportCommon::checkInvalidSheetName(importXlsx);
}
