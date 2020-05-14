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

const QVector<std::pair<QString, QString>> ImportXlsxTest::sheets_{
    {"Sheet1", "xl/worksheets/sheet1.xml"},
    {"Sheet2", "xl/worksheets/sheet2.xml"},
    {"Sheet3(empty)", "xl/worksheets/sheet3.xml"},
    {"Sheet4", "xl/worksheets/sheet4.xml"},
    {"Sheet5", "xl/worksheets/sheet5.xml"},
    {"Sheet6", "xl/worksheets/sheet6.xml"},
    {"Sheet7", "xl/worksheets/sheet7.xml"},
    {"Sheet9", "xl/worksheets/sheet8.xml"},
    {"testAccounts", "xl/worksheets/sheet9.xml"}};

const QStringList ImportXlsxTest::sharedStrings_{
    QStringLiteral("Text"),
    QStringLiteral("Numeric"),
    QStringLiteral("Date"),
    QStringLiteral("Item 0, 0"),
    QStringLiteral("Other item"),
    QStringLiteral("Trait #1"),
    QStringLiteral("Value #1"),
    QStringLiteral("Transaction date"),
    QStringLiteral("Units"),
    QStringLiteral("Price"),
    QStringLiteral("Price per unit"),
    QStringLiteral("Another trait"),
    QStringLiteral("trait 1"),
    QStringLiteral("brown"),
    QStringLiteral("trait 2"),
    QStringLiteral("red"),
    QStringLiteral("trait 3"),
    QStringLiteral("yellow"),
    QStringLiteral("trait 4"),
    QStringLiteral("black"),
    QStringLiteral("trait 16"),
    QStringLiteral("trait 17"),
    QStringLiteral("trait 18"),
    QStringLiteral("white"),
    QStringLiteral("trait 19"),
    QStringLiteral("trait 20"),
    QStringLiteral("trait 21"),
    QStringLiteral("pink"),
    QStringLiteral("trait 22"),
    QStringLiteral("blue"),
    QStringLiteral("trait 5"),
    QStringLiteral("trait 6"),
    QStringLiteral("trait 7"),
    QStringLiteral("trait 8"),
    QStringLiteral("cena nier"),
    QStringLiteral("pow"),
    QStringLiteral("cena metra"),
    QStringLiteral("data transakcji"),
    QStringLiteral("text"),
    QStringLiteral("a"),
    QStringLiteral("b"),
    QStringLiteral("c"),
    QStringLiteral("name"),
    QStringLiteral("date"),
    QStringLiteral("mass (kg)"),
    QStringLiteral("height"),
    QStringLiteral("Sx"),
    QStringLiteral("Sy"),
    QStringLiteral("Sxx"),
    QStringLiteral("Syy"),
    QStringLiteral("Sxy"),
    QStringLiteral("n"),
    QStringLiteral("a (beta w wiki)"),
    QStringLiteral("b (alfa w wiki)"),
    QStringLiteral("y=ax+b"),
    QStringLiteral("modificator"),
    QStringLiteral("x"),
    QStringLiteral("y"),
    QStringLiteral("Pow"),
    QStringLiteral("Cena"),
    QStringLiteral("cena_m"),
    QStringLiteral("Średnia:"),
    QStringLiteral("Mediana:"),
    QStringLiteral("Q10"),
    QStringLiteral("Q25"),
    QStringLiteral("Q50"),
    QStringLiteral("Q75"),
    QStringLiteral("Q90"),
    QStringLiteral("second"),
    QStringLiteral("third"),
    QStringLiteral("fourth"),
    QStringLiteral("s"),
    QStringLiteral("d"),
    QStringLiteral("user"),
    QStringLiteral("pass"),
    QStringLiteral("lic_exp"),
    QStringLiteral("uwagi"),
    QStringLiteral("test1"),
    QStringLiteral("test123"),
    QStringLiteral("zawsze aktualna"),
    QStringLiteral("test2"),
    QStringLiteral("konto zablokowane"),
    QStringLiteral("test3")};

const QList<int> ImportXlsxTest::dateStyles_{14,  15,  16,  17,  22, 165,
                                             167, 170, 171, 172, 173};
const QList<int> ImportXlsxTest::allStyles_{
    164, 164, 165, 164, 166, 167, 168, 169, 169, 170, 164, 164, 171, 172, 173};

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
    QCOMPARE(sheets_.size(), ImportCommon::getSheetNames().size());
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
    QTest::newRow(("Limited data to 10 in " + sheetName).toStdString().c_str())
        << sheetName << 10u << sheetData;
    QTest::newRow(("Limited data to 2 in " + sheetName).toStdString().c_str())
        << sheetName << 2u << sheetData;
    sheetName = sheets_[5].first;
    QVector<QVector<QVariant>> expectedValues;
    const unsigned int rowLimit{12u};
    sheetData = ImportCommon::getDataForSheet(sheetName);
    for (unsigned int i = 0; i < rowLimit; ++i)
        expectedValues.append(sheetData[static_cast<int>(i)]);
    expectedValues = convertDataToUseSharedStrings(expectedValues);
    QTest::newRow(
        ("Limited data to " + QString::number(rowLimit) + " in " + sheetName)
            .toStdString()
            .c_str())
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
