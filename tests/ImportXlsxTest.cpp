#include "ImportXlsxTest.h"

#include <eible/ExportXlsx.h>
#include <eible/ImportXlsx.h>
#include <QBuffer>
#include <QTableView>
#include <QTest>

#include "ImportCommon.h"
#include "TestTableModel.h"

ImportXlsxTest::ImportXlsxTest(QObject* parent) : QObject(parent) {}

void ImportXlsxTest::testRetrievingSheetNames() const
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

void ImportXlsxTest::testGetDateStyles() const
{
    QFile xlsxTestFile(testFileName_);
    ImportXlsx importXlsx(xlsxTestFile);
    auto [success, actualDateStyle]{importXlsx.getDateStyles()};
    QCOMPARE(success, true);
    QCOMPARE(actualDateStyle, dateStyles_);
}

void ImportXlsxTest::testGetDateStylesNoContent() const
{
    QFile xlsxTestFile(templateFileName_);
    ImportXlsx importXlsx(xlsxTestFile);
    auto [success, actualDateStyle]{importXlsx.getDateStyles()};
    QCOMPARE(success, true);
    QCOMPARE(actualDateStyle, QList({14, 15, 16, 17, 22}));
}

void ImportXlsxTest::testGetAllStyles() const
{
    QFile xlsxTestFile(testFileName_);
    ImportXlsx importXlsx(xlsxTestFile);
    auto [success, actualAllStyle]{importXlsx.getAllStyles()};
    QCOMPARE(success, true);
    QCOMPARE(actualAllStyle, allStyles_);
}

void ImportXlsxTest::testGetAllStylesNoContent() const
{
    QFile xlsxTestFile(templateFileName_);
    ImportXlsx importXlsx(xlsxTestFile);
    auto [success, actualAllStyle]{importXlsx.getAllStyles()};
    QCOMPARE(success, true);
    QCOMPARE(actualAllStyle, QList({0, 164, 10, 14, 4, 0, 0, 3}));
}

void ImportXlsxTest::testGetSharedStrings() const
{
    QFile xlsxTestFile(testFileName_);
    ImportXlsx importXlsx(xlsxTestFile);
    auto [success, actualSharedStrings]{importXlsx.getSharedStrings()};
    QCOMPARE(success, true);
    QCOMPARE(actualSharedStrings, sharedStrings_);
}

void ImportXlsxTest::testGetSharedStringsNoContent() const
{
    QFile xlsxTestFile(templateFileName_);
    ImportXlsx importXlsx(xlsxTestFile);
    auto [success, actualSharedStrings]{importXlsx.getSharedStrings()};
    QCOMPARE(success, true);
    QCOMPARE(actualSharedStrings, {});
}

void ImportXlsxTest::testGetColumnList_data()
{
    ImportCommon::prepareDataForGetColumnListTest();
}

void ImportXlsxTest::testGetColumnList() const
{
    QFile testFile(testFileName_);
    ImportXlsx importXlsx(testFile);
    ImportCommon::checkGetColumnList(importXlsx);
}

void ImportXlsxTest::testSettingEmptyColumnName() const
{
    QFile testFile(testFileName_);
    ImportXlsx importXlsx(testFile);
    ImportCommon::checkSettingEmptyColumnName(importXlsx);
}

void ImportXlsxTest::testGetColumnListTwoSheets() const
{
    QFile testFile(testFileName_);
    ImportXlsx importXlsx(testFile);
    ImportCommon::checkGetColumnListTwoSheets(importXlsx);
}

void ImportXlsxTest::testGetColumnTypes_data()
{
    ImportCommon::prepareDataForGetColumnTypes();
}

void ImportXlsxTest::testGetColumnTypes() const
{
    QFile testFile(testFileName_);
    ImportXlsx importXlsx(testFile);
    ImportCommon::checkGetColumnTypes(importXlsx);
}

void ImportXlsxTest::testGetColumnCount_data()
{
    ImportCommon::prepareDataForGetColumnCountTest();
}

void ImportXlsxTest::testGetColumnCount() const
{
    QFile testFile(testFileName_);
    ImportXlsx importXlsx(testFile);
    ImportCommon::checkGetColumnCount(importXlsx);
}

void ImportXlsxTest::testGetRowCount_data()
{
    ImportCommon::prepareDataForGetRowCountTest();
}

void ImportXlsxTest::testGetRowCount() const
{
    QFile testFile(testFileName_);
    ImportXlsx importXlsx(testFile);
    ImportCommon::checkGetRowCount(importXlsx);
}

void ImportXlsxTest::testGetRowAndColumnCountViaGetColumnTypes_data()
{
    ImportCommon::prepareDataForGetRowAndColumnCountViaGetColumnTypes();
}

void ImportXlsxTest::testGetRowAndColumnCountViaGetColumnTypes() const
{
    QFile testFile(testFileName_);
    ImportXlsx importXlsx(testFile);
    ImportCommon::testGetRowAndColumnCountViaGetColumnTypes(importXlsx);
}

QVector<QVector<QVariant>> ImportXlsxTest::convertDataToUseSharedStrings(
    const QVector<QVector<QVariant>>& inputData) const
{
    QVector<QVector<QVariant>> outputData{inputData};
    for (auto& row : outputData)
        for (auto& item : row)
            if (item.typeId() == QMetaType::QString && !item.isNull())
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

void ImportXlsxTest::testGetData() const
{
    QFile testFile(testFileName_);
    ImportXlsx importXlsx(testFile);
    ImportCommon::checkGetData(importXlsx);
}

void ImportXlsxTest::testGetDataLimitRows_data()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<int>("rowLimit");
    QTest::addColumn<QVector<QVector<QVariant>>>("expectedData");

    QString sheetName{sheets_[0].first};
    QVector<QVector<QVariant>> sheetData{convertDataToUseSharedStrings(
        ImportCommon::getDataForSheet(sheetName))};
    int rowLimit{10};
    QString testName{"Limited data to " + QString::number(rowLimit) + " in " +
                     sheetName};
    QTest::newRow(testName.toStdString().c_str())
        << sheetName << rowLimit << sheetData;

    rowLimit = 2;
    testName =
        "Limited data to " + QString::number(rowLimit) + " in " + sheetName;
    QTest::newRow(testName.toStdString().c_str())
        << sheetName << rowLimit << sheetData;

    sheetName = sheets_[5].first;
    rowLimit = 12;
    QVector<QVector<QVariant>> expectedValues;
    expectedValues.reserve(rowLimit);
    sheetData = ImportCommon::getDataForSheet(sheetName);
    for (int i{0}; i < rowLimit; ++i)
        expectedValues.append(sheetData[i]);
    expectedValues = convertDataToUseSharedStrings(expectedValues);
    testName =
        "Limited data to " + QString::number(rowLimit) + " in " + sheetName;
    QTest::newRow(testName.toStdString().c_str())
        << sheetName << rowLimit << expectedValues;
}

void ImportXlsxTest::testGetDataLimitRows() const
{
    QFile testFile(testFileName_);
    ImportXlsx importXlsx(testFile);
    ImportCommon::checkGetDataLimitRows(importXlsx);
}

void ImportXlsxTest::testGetDataExcludeColumns_data()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<QVector<int>>("excludedColumns");
    QTest::addColumn<QVector<QVector<QVariant>>>("expectedData");

    const QString sheetName{sheets_[0].first};
    QString testName{"Get data with excluded column 1 in " + sheetName};
    const QVector<QVector<QVariant>> sheetData{convertDataToUseSharedStrings(
        ImportCommon::getDataForSheet(sheetName))};
    QTest::newRow(qUtf8Printable(testName))
        << sheetName << QVector<int>{1}
        << ImportCommon::getDataWithoutColumns(sheetData, {1});

    testName = "Get data with excluded column 0 and 2 in " + sheetName;
    QTest::newRow(qUtf8Printable(testName))
        << sheetName << QVector<int>{0, 2}
        << ImportCommon::getDataWithoutColumns(sheetData, {2, 0});
}

void ImportXlsxTest::testGetDataExcludeColumns() const
{
    QFile testFile(testFileName_);
    ImportXlsx importXlsx(testFile);
    ImportCommon::checkGetDataExcludeColumns(importXlsx);
}

void ImportXlsxTest::testGetDataExcludeInvalidColumn() const
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
    auto [success, sheetNames]{importXlsx.getSheetNames()};
    QCOMPARE(success, true);
    QCOMPARE(sheetNames.size(), 1);
    QBENCHMARK { importXlsx.getData(sheetNames.front(), {}); }
}

void ImportXlsxTest::testEmittingProgressPercentChangedEmptyFile() const
{
    QFile testFile(templateFileName_);
    ImportXlsx importXlsx(testFile);
    ImportCommon::checkEmittingProgressPercentChangedEmptyFile(importXlsx);
}

void ImportXlsxTest::testEmittingProgressPercentChangedSmallFile() const
{
    QFile testFile(testFileName_);
    ImportXlsx importXlsx(testFile);
    ImportCommon::checkEmittingProgressPercentChangedSmallFile(importXlsx);
}

void ImportXlsxTest::testEmittingProgressPercentChangedBigFile()
{
    QFile testFile(QStringLiteral(":/res/mediumFile.xlsx"));
    ImportXlsx importXlsx(testFile);
    ImportCommon::checkEmittingProgressPercentChangedBigFile(importXlsx);
}

void ImportXlsxTest::testInvalidSheetName() const
{
    QFile testFile(testFileName_);
    ImportXlsx importXlsx(testFile);
    ImportCommon::checkInvalidSheetName(importXlsx);
}

void ImportXlsxTest::testDamagedFile()
{
    const QString sheet{QString::fromLatin1("mySheet")};
    QFile testFile(QString::fromLatin1(":/res/testXlsx_damaged.xlsx"));
    ImportXlsx importXlsx(testFile);
    QCOMPARE(importXlsx.getSheetNames().first, false);
    QCOMPARE(importXlsx.getColumnTypes(sheet).first, false);
    QCOMPARE(importXlsx.getColumnCount(sheet).first, false);
    QCOMPARE(importXlsx.getColumnNames(sheet).first, false);
    QCOMPARE(importXlsx.getSharedStrings().first, false);
    QCOMPARE(importXlsx.getDateStyles().first, false);
    QCOMPARE(importXlsx.getAllStyles().first, false);
    QCOMPARE(importXlsx.getLimitedData(sheet, {}, 10).first, false);
    QCOMPARE(importXlsx.getData(sheet, {}).first, false);
    QCOMPARE(importXlsx.getRowCount(sheet).first, false);
}
