#include "ImportXlsxTest.h"

#include <QBuffer>
#include <QTest>

#include "ImportXlsx.h"

QList<std::pair<QString, QString>> ImportXlsxTest::sheets_{
    {"Sheet1", "xl/worksheets/sheet1.xml"},
    {"Sheet2", "xl/worksheets/sheet2.xml"},
    {"Sheet3(empty)", "xl/worksheets/sheet3.xml"},
    {"Sheet4", "xl/worksheets/sheet4.xml"},
    {"Sheet5", "xl/worksheets/sheet5.xml"},
    {"Sheet6", "xl/worksheets/sheet6.xml"},
    {"Sheet7", "xl/worksheets/sheet7.xml"}};

QStringList ImportXlsxTest::sharedStrings_{"Text",
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
                                           "Q90"};

QList<int> ImportXlsxTest::dateStyles_{14,  15,  16,  17,  22,
                                       165, 167, 170, 171, 172};
QList<int> ImportXlsxTest::allStyles_{164, 164, 165, 164, 166, 167, 168,
                                      169, 170, 164, 164, 171, 172};

QStringList ImportXlsxTest::testSheet1Columns_ = {"Text", "Numeric", "Date"};
QStringList ImportXlsxTest::testSheet2Columns_ = {
    "Trait #1", "Value #1",       "Transaction date", "Units",
    "Price",    "Price per unit", "Another trait"};
QStringList ImportXlsxTest::testSheet3Columns_ = {};
QStringList ImportXlsxTest::testSheet4Columns_ = {
    "cena nier", "pow", "cena metra", "data transakcji", "text"};
QStringList ImportXlsxTest::testSheet5Columns_ = {
    "name", "date", "mass (kg)", "height", "---", "---",
    "---",  "---",  "---",       "---",    "---", "---"};
QStringList ImportXlsxTest::testSheet6Columns_ = {"modificator", "x", "y"};
QStringList ImportXlsxTest::testSheet7Columns_ = {"Pow", "Cena", "cena_m"};

void ImportXlsxTest::testRetrievingSheetNames()
{
    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    auto [success, actualNames] = importXlsx.getSheetNames();
    QCOMPARE(success, true);
    QStringList sheets;
    for (const auto& [sheetName, sheetPath] : sheets_)
        sheets << sheetName;
    QCOMPARE(actualNames, sheets);
}

void ImportXlsxTest::testRetrievingSheetNamesFromEmptyFile()
{
    QByteArray byteArray;
    QBuffer emptyBuffer(&byteArray);
    ImportXlsx importXlsx(emptyBuffer);
    auto [success, actualNames] = importXlsx.getSheetNames();
    QCOMPARE(success, false);
    QCOMPARE(actualNames, {});
}

void ImportXlsxTest::testGetDateStyles()
{
    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    auto [success, actualDateStyle] = importXlsx.getDateStyles();
    QCOMPARE(success, true);
    QCOMPARE(actualDateStyle, dateStyles_);
}

void ImportXlsxTest::testGetDateStylesNoContent()
{
    QFile xlsxTestFile(QStringLiteral(":/template.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    auto [success, actualDateStyle] = importXlsx.getDateStyles();
    QCOMPARE(success, true);
    QCOMPARE(actualDateStyle, QList({14, 15, 16, 17, 22}));
}

void ImportXlsxTest::testGetAllStyles()
{
    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    auto [success, actualAllStyle] = importXlsx.getAllStyles();
    QCOMPARE(success, true);
    QCOMPARE(actualAllStyle, allStyles_);
}

void ImportXlsxTest::testGetAllStylesNoContent()
{
    QFile xlsxTestFile(QStringLiteral(":/template.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    auto [success, actualAllStyle] = importXlsx.getAllStyles();
    QCOMPARE(success, true);
    QCOMPARE(actualAllStyle, QList({0, 164, 10, 14, 4, 0, 0, 3}));
}

void ImportXlsxTest::testGetSharedStrings()
{
    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    auto [success, actualSharedStrings] = importXlsx.getSharedStrings();
    QCOMPARE(success, true);
    QCOMPARE(actualSharedStrings, sharedStrings_);
}

void ImportXlsxTest::testGetSharedStringsNoContent()
{
    QFile xlsxTestFile(QStringLiteral(":/template.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    auto [success, actualSharedStrings] = importXlsx.getSharedStrings();
    QCOMPARE(success, true);
    QCOMPARE(actualSharedStrings, {});
}

void ImportXlsxTest::testGetColumnList_data()
{
    QTest::addColumn<QString>("sheetPath");
    QTest::addColumn<QStringList>("expectedColumnList");
    QTest::newRow("Columns in Sheet1") << "Sheet1" << testSheet1Columns_;
    QTest::newRow("Columns in Sheet2") << "Sheet2" << testSheet2Columns_;
    QTest::newRow("Columns in Sheet3(empty)")
        << "Sheet3(empty)" << testSheet3Columns_;
    QTest::newRow("Columns in Sheet4") << "Sheet4" << testSheet4Columns_;
    QTest::newRow("Columns in Sheet5") << "Sheet5" << testSheet5Columns_;
    QTest::newRow("Columns in Sheet6") << "Sheet6" << testSheet6Columns_;
    QTest::newRow("Columns in Sheet7") << "Sheet7" << testSheet7Columns_;
}

void ImportXlsxTest::testGetColumnList()
{
    QFETCH(QString, sheetPath);
    QFETCH(QStringList, expectedColumnList);
    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    importXlsx.setSharedStrings(sharedStrings_);
    importXlsx.setSheets(sheets_);
    auto [success, actualColumnList] = importXlsx.getColumnList(sheetPath);
    QCOMPARE(success, true);
    QCOMPARE(actualColumnList, expectedColumnList);
}

void ImportXlsxTest::testSettingEmptyColumnName()
{
    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    const QString newEmptyColumnName{"<empty column>"};
    importXlsx.setNameForEmptyColumn(newEmptyColumnName);
    importXlsx.setSharedStrings(sharedStrings_);
    importXlsx.setSheets(sheets_);
    auto [success, actualColumnList] = importXlsx.getColumnList("Sheet5");

    std::list<QString> expectedColumnList(testSheet5Columns_.size());
    std::replace_copy(testSheet5Columns_.begin(), testSheet5Columns_.end(),
                      expectedColumnList.begin(), QString("---"),
                      newEmptyColumnName);

    QCOMPARE(success, true);
    QCOMPARE(actualColumnList, QStringList::fromStdList(expectedColumnList));
}

void ImportXlsxTest::testGetColumnListTwoSheets()
{
    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    importXlsx.setSharedStrings(sharedStrings_);
    importXlsx.setSheets(sheets_);
    auto [success, actualColumnList] = importXlsx.getColumnList("Sheet5");
    QCOMPARE(success, true);
    QCOMPARE(actualColumnList, testSheet5Columns_);

    std::tie(success, actualColumnList) = importXlsx.getColumnList("Sheet1");
    QCOMPARE(success, true);
    QCOMPARE(actualColumnList, testSheet1Columns_);
}

void ImportXlsxTest::testGetColumnTypes_data()
{
    QTest::addColumn<QString>("sheetPath");
    QTest::addColumn<QVector<ColumnType>>("expectedColumnTypes");

    QTest::newRow("Column types in Sheet1")
        << "Sheet1"
        << QVector({ColumnType::STRING, ColumnType::NUMBER, ColumnType::DATE});
    QTest::newRow("Column types in Sheet2")
        << "Sheet2"
        << QVector({ColumnType::STRING, ColumnType::NUMBER, ColumnType::DATE,
                    ColumnType::NUMBER, ColumnType::NUMBER, ColumnType::NUMBER,
                    ColumnType::STRING});
    QTest::newRow("Column types in Sheet3(empty)")
        << "Sheet3(empty)" << QVector<ColumnType>();
    QTest::newRow("Column types in Sheet4")
        << "Sheet4"
        << QVector({ColumnType::NUMBER, ColumnType::NUMBER, ColumnType::NUMBER,
                    ColumnType::DATE, ColumnType::STRING});
    QTest::newRow("Column types in Sheet5")
        << "Sheet5"
        << QVector({ColumnType::STRING, ColumnType::STRING, ColumnType::NUMBER,
                    ColumnType::NUMBER, ColumnType::STRING, ColumnType::NUMBER,
                    ColumnType::STRING, ColumnType::STRING, ColumnType::NUMBER,
                    ColumnType::NUMBER, ColumnType::NUMBER,
                    ColumnType::NUMBER});
    QTest::newRow("Column types in Sheet6")
        << "Sheet6"
        << QVector(
               {ColumnType::NUMBER, ColumnType::NUMBER, ColumnType::NUMBER});
    QTest::newRow("Column types in Sheet7")
        << "Sheet7"
        << QVector({ColumnType::NUMBER, ColumnType::NUMBER, ColumnType::NUMBER,
                    ColumnType::STRING, ColumnType::STRING,
                    ColumnType::NUMBER});
}

void ImportXlsxTest::testGetColumnTypes()
{
    QFETCH(QString, sheetPath);
    QFETCH(QVector<ColumnType>, expectedColumnTypes);

    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    importXlsx.setSharedStrings(sharedStrings_);
    importXlsx.setDateStyles(dateStyles_);
    importXlsx.setAllStyles(allStyles_);
    importXlsx.setSheets(sheets_);

    auto [success, columnTypes] = importXlsx.getColumnTypes(sheetPath);
    QCOMPARE(success, true);
    QCOMPARE(columnTypes, expectedColumnTypes);
}
