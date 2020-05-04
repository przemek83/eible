#include "ImportXlsxTest.h"

#include <ExportXlsx.h>
#include <ImportXlsx.h>
#include <QBuffer>
#include <QSignalSpy>
#include <QTableView>
#include <QTest>

#include "ImportCommon.h"
#include "TestTableModel.h"

const QString ImportXlsxTest::testFileName_{QStringLiteral(":/testXlsx.xlsx")};
const QString ImportXlsxTest::templateFileName_{
    QStringLiteral(":/template.xlsx")};

const QList<std::pair<QString, QString>> ImportXlsxTest::sheets_{
    {"Sheet1", "xl/worksheets/sheet1.xml"},
    {"Sheet2", "xl/worksheets/sheet2.xml"},
    {"Sheet3(empty)", "xl/worksheets/sheet3.xml"},
    {"Sheet4", "xl/worksheets/sheet4.xml"},
    {"Sheet5", "xl/worksheets/sheet5.xml"},
    {"Sheet6", "xl/worksheets/sheet6.xml"},
    {"Sheet7", "xl/worksheets/sheet7.xml"},
    {"testAccounts", "xl/worksheets/sheet8.xml"}};

const QStringList ImportXlsxTest::sharedStrings_{"Text",
                                                 "Numeric",
                                                 "Date",
                                                 "Item 0, 0",
                                                 "Other item",
                                                 "Trait #1",
                                                 "Value #1",
                                                 "Transaction date",
                                                 "Units",
                                                 "Price",
                                                 "Price per unit",
                                                 "Another trait",
                                                 "trait 1",
                                                 "brown",
                                                 "trait 2",
                                                 "red",
                                                 "trait 3",
                                                 "yellow",
                                                 "trait 4",
                                                 "black",
                                                 "trait 16",
                                                 "trait 17",
                                                 "trait 18",
                                                 "white",
                                                 "trait 19",
                                                 "trait 20",
                                                 "trait 21",
                                                 "pink",
                                                 "trait 22",
                                                 "blue",
                                                 "trait 5",
                                                 "trait 6",
                                                 "trait 7",
                                                 "trait 8",
                                                 "cena nier",
                                                 "pow",
                                                 "cena metra",
                                                 "data transakcji",
                                                 "text",
                                                 "a",
                                                 "b",
                                                 "c",
                                                 "name",
                                                 "date",
                                                 "mass (kg)",
                                                 "height",
                                                 "Sx",
                                                 "Sy",
                                                 "Sxx",
                                                 "Syy",
                                                 "Sxy",
                                                 "n",
                                                 "a (beta w wiki)",
                                                 "b (alfa w wiki)",
                                                 "y=ax+b",
                                                 "modificator",
                                                 "x",
                                                 "y",
                                                 "Pow",
                                                 "Cena",
                                                 "cena_m",
                                                 "Średnia:",
                                                 "Mediana:",
                                                 "Q10",
                                                 "Q25",
                                                 "Q50",
                                                 "Q75",
                                                 "Q90",
                                                 "user",
                                                 "pass",
                                                 "lic_exp",
                                                 "uwagi",
                                                 "test1",
                                                 "test123",
                                                 "zawsze aktualna",
                                                 "test2",
                                                 "konto zablokowane",
                                                 "test3"};

const QList<int> ImportXlsxTest::dateStyles_{14,  15,  16,  17,  22,
                                             165, 167, 170, 171, 172};
const QList<int> ImportXlsxTest::allStyles_{164, 164, 165, 164, 166, 167, 168,
                                            169, 169, 170, 164, 164, 171, 172};

void ImportXlsxTest::testRetrievingSheetNames()
{
    ImportCommon::checkRetrievingSheetNames<ImportXlsx>(testFileName_);
}

void ImportXlsxTest::testRetrievingSheetNamesFromEmptyFile()
{
    ImportCommon::checkRetrievingSheetNamesFromEmptyFile<ImportXlsx>();
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
    ImportCommon::checkGetColumnList<ImportXlsx>(testFileName_);
}

void ImportXlsxTest::testSettingEmptyColumnName()
{
    ImportCommon::checkSettingEmptyColumnName<ImportXlsx>(testFileName_);
}

void ImportXlsxTest::testGetColumnListTwoSheets()
{
    ImportCommon::checkGetColumnListTwoSheets<ImportXlsx>(testFileName_);
}

void ImportXlsxTest::testGetColumnTypes_data()
{
    ImportCommon::prepareDataForGetColumnTypes();
}

void ImportXlsxTest::testGetColumnTypes()
{
    ImportCommon::checkGetColumnTypes<ImportXlsx>(testFileName_);
}

void ImportXlsxTest::testGetColumnCount_data()
{
    ImportCommon::prepareDataForGetColumnCountTest();
}

void ImportXlsxTest::testGetColumnCount()
{
    ImportCommon::checkGetColumnCount<ImportXlsx>(testFileName_);
}

void ImportXlsxTest::testGetRowCount_data()
{
    ImportCommon::prepareDataForGetRowCountTest();
}

void ImportXlsxTest::testGetRowCount()
{
    ImportCommon::checkGetRowCount<ImportXlsx>(testFileName_);
}

void ImportXlsxTest::testGetRowAndColumnCountViaGetColumnTypes_data()
{
    ImportCommon::prepareDataForGetRowAndColumnCountViaGetColumnTypes();
}

void ImportXlsxTest::testGetRowAndColumnCountViaGetColumnTypes()
{
    ImportCommon::testGetRowAndColumnCountViaGetColumnTypes<ImportXlsx>(
        testFileName_);
}

QVector<QVector<QVariant>> ImportXlsxTest::convertDataToUseSharedStrings(
    QVector<QVector<QVariant>> inputData)
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
    for (int i = 0; i < sheets_.size(); ++i)
    {
        const QString& sheetName{sheets_[i].first};
        QVector<QVector<QVariant>> sheetData{
            ImportCommon::getDataForSheet(sheetName)};
        sheetData = convertDataToUseSharedStrings(sheetData);
        QTest::newRow(("Data in " + sheetName).toStdString().c_str())
            << sheetName << sheetData;
    }
}

void ImportXlsxTest::testGetData()
{
    ImportCommon::checkGetData<ImportXlsx>(testFileName_);
}

void ImportXlsxTest::testGetDataLimitRows_data()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<unsigned int>("rowLimit");
    QTest::addColumn<QVector<QVector<QVariant>>>("expectedData");
    QString sheetName{sheets_[0].first};
    QVector<QVector<QVariant>> sheetData{convertDataToUseSharedStrings(
        ImportCommon::getDataForSheet(sheetName))};
    QTest::newRow(("Limited data to 10 in " + sheetName).toStdString().c_str())
        << sheetName << 10u << sheetData;
    QTest::newRow(("Limited data to 2 in " + sheetName).toStdString().c_str())
        << sheetName << 2u << sheetData;
    sheetName = sheets_[5].first;
    QVector<QVector<QVariant>> expectedValues;
    const unsigned int rowLimit{12u};
    sheetData = ImportCommon::getDataForSheet(sheetName);
    for (unsigned int i = 0; i < rowLimit; ++i)
        expectedValues.append(sheetData[i]);
    expectedValues = convertDataToUseSharedStrings(expectedValues);
    QTest::newRow(
        ("Limited data to " + QString::number(rowLimit) + " in " + sheetName)
            .toStdString()
            .c_str())
        << sheetName << rowLimit << expectedValues;
}

void ImportXlsxTest::testGetDataLimitRows()
{
    ImportCommon::checkGetDataLimitRows<ImportXlsx>(testFileName_);
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
    ImportCommon::checkGetDataExcludeColumns<ImportXlsx>(testFileName_);
}

void ImportXlsxTest::testGetDataExcludeInvalidColumn()
{
    ImportCommon::checkGetDataExcludeInvalidColumn<ImportXlsx>(testFileName_);
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
    ImportCommon::checkEmittingProgressPercentChangedEmptyFile<ImportXlsx>(
        templateFileName_);
}

void ImportXlsxTest::testEmittingProgressPercentChangedSmallFile()
{
    ImportCommon::checkEmittingProgressPercentChangedSmallFile<ImportXlsx>(
        testFileName_);
}

void ImportXlsxTest::testEmittingProgressPercentChangedBigFile()
{
    ImportCommon::checkEmittingProgressPercentChangedBigFile<ImportXlsx>(
        ":/mediumFile.xlsx");
}
