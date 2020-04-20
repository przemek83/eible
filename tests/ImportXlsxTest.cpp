#include "ImportXlsxTest.h"

#include <QBuffer>
#include <QTest>

#include "ImportXlsx.h"

QMap<QString, QString> ImportXlsxTest::sheetNames_{
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

QList<int> ImportXlsxTest::dateStyle_{14,  15,  16,  17,  22,
                                      165, 167, 170, 171, 172};
QList<int> ImportXlsxTest::allStyles_{164, 164, 165, 167, 168, 164, 169,
                                      170, 166, 164, 164, 171, 172};

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
    auto [success, actualNames] = importXlsx.getSheetList();
    QCOMPARE(success, true);
    QCOMPARE(actualNames, sheetNames_);
}

void ImportXlsxTest::testRetrievingSheetNamesFromEmptyFile()
{
    QByteArray byteArray;
    QBuffer emptyBuffer(&byteArray);
    ImportXlsx importXlsx(emptyBuffer);
    auto [success, actualNames] = importXlsx.getSheetList();
    QCOMPARE(success, false);
    QCOMPARE(actualNames, {});
}

void ImportXlsxTest::testGetStyles()
{
    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    auto [success, actualDateStyle, actualAllStyles] = importXlsx.getStyles();
    QCOMPARE(success, true);
    QCOMPARE(actualDateStyle, dateStyle_);
    QCOMPARE(actualAllStyles, allStyles_);
}

void ImportXlsxTest::testGetStylesNoContent()
{
    QFile xlsxTestFile(QStringLiteral(":/template.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    auto [success, dateStyle, allStyles] = importXlsx.getStyles();
    QCOMPARE(success, true);
    QCOMPARE(dateStyle, QList({14, 15, 16, 17, 22}));
    QCOMPARE(allStyles, QList({0, 164, 10, 14, 4, 0, 0, 3}));
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
    QTest::newRow("Columns in Sheet1")
        << sheetNames_["Sheet1"] << testSheet1Columns_;
    QTest::newRow("Columns in Sheet2")
        << sheetNames_["Sheet2"] << testSheet2Columns_;
    QTest::newRow("Columns in Sheet3(empty)")
        << sheetNames_["Sheet3(empty)"] << testSheet3Columns_;
    QTest::newRow("Columns in Sheet4")
        << sheetNames_["Sheet4"] << testSheet4Columns_;
    QTest::newRow("Columns in Sheet5")
        << sheetNames_["Sheet5"] << testSheet5Columns_;
    QTest::newRow("Columns in Sheet6")
        << sheetNames_["Sheet6"] << testSheet6Columns_;
    QTest::newRow("Columns in Sheet7")
        << sheetNames_["Sheet7"] << testSheet7Columns_;
}

void ImportXlsxTest::testGetColumnList()
{
    QFETCH(QString, sheetPath);
    QFETCH(QStringList, expectedColumnList);
    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    QHash<QString, int> sharedStringsMap =
        createSharedStringsMap(sharedStrings_);
    auto [success, actualColumnList] =
        importXlsx.getColumnList(sheetPath, sharedStringsMap);
    QCOMPARE(success, true);
    QCOMPARE(actualColumnList, expectedColumnList);
}

void ImportXlsxTest::testSettingEmptyColumnName()
{
    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    const QString newEmptyColumnName{"<empty column>"};
    importXlsx.setNameForEmptyColumn(newEmptyColumnName);
    QHash<QString, int> sharedStringsMap =
        createSharedStringsMap(sharedStrings_);
    auto [success, actualColumnList] =
        importXlsx.getColumnList(sheetNames_["Sheet5"], sharedStringsMap);

    std::list<QString> expectedColumnList(testSheet5Columns_.size());
    std::replace_copy(testSheet5Columns_.begin(), testSheet5Columns_.end(),
                      expectedColumnList.begin(), QString("---"),
                      newEmptyColumnName);

    QCOMPARE(success, true);
    QCOMPARE(actualColumnList, QStringList::fromStdList(expectedColumnList));
}

void ImportXlsxTest::testGetColumnTypes_data()
{
    QTest::addColumn<QString>("sheetPath");
    QTest::addColumn<int>("columnCount");
    QTest::addColumn<QVector<ColumnType>>("expectedColumnTypes");

    QTest::newRow("Column types in Sheet1")
        << sheetNames_["Sheet1"] << 3
        << QVector({ColumnType::STRING, ColumnType::NUMBER, ColumnType::DATE});
    QTest::newRow("Column types in Sheet2")
        << sheetNames_["Sheet2"] << 7
        << QVector({ColumnType::STRING, ColumnType::NUMBER, ColumnType::DATE,
                    ColumnType::NUMBER, ColumnType::NUMBER, ColumnType::NUMBER,
                    ColumnType::STRING});
    QTest::newRow("Column types in Sheet3(empty)")
        << sheetNames_["Sheet3(empty)"] << 0 << QVector<ColumnType>();
    QTest::newRow("Column types in Sheet4")
        << sheetNames_["Sheet4"] << 5
        << QVector({ColumnType::NUMBER, ColumnType::NUMBER, ColumnType::NUMBER,
                    ColumnType::DATE, ColumnType::STRING});
    QTest::newRow("Column types in Sheet5")
        << sheetNames_["Sheet5"] << 12
        << QVector({ColumnType::STRING, ColumnType::STRING, ColumnType::NUMBER,
                    ColumnType::NUMBER, ColumnType::STRING, ColumnType::NUMBER,
                    ColumnType::STRING, ColumnType::STRING, ColumnType::NUMBER,
                    ColumnType::NUMBER, ColumnType::NUMBER,
                    ColumnType::NUMBER});
    QTest::newRow("Column types in Sheet6")
        << sheetNames_["Sheet6"] << 3
        << QVector(
               {ColumnType::NUMBER, ColumnType::NUMBER, ColumnType::NUMBER});
    QTest::newRow("Column types in Sheet7")
        << sheetNames_["Sheet7"] << 6
        << QVector({ColumnType::NUMBER, ColumnType::NUMBER, ColumnType::NUMBER,
                    ColumnType::STRING, ColumnType::STRING,
                    ColumnType::NUMBER});
}

void ImportXlsxTest::testGetColumnTypes()
{
    QFETCH(QString, sheetPath);
    QFETCH(int, columnCount);
    QFETCH(QVector<ColumnType>, expectedColumnTypes);

    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    QHash<QString, int> sharedStringsMap =
        createSharedStringsMap(sharedStrings_);

    auto [success, columnTypes] = importXlsx.getColumnTypes(
        sheetPath, columnCount, sharedStringsMap, dateStyle_, allStyles_);
    QCOMPARE(success, true);
    QCOMPARE(columnTypes, expectedColumnTypes);
}

QHash<QString, int> ImportXlsxTest::createSharedStringsMap(
    const QStringList& sharedStrings)
{
    QHash<QString, int> sharedStringsMap;
    int currentSharedStringIndex{0};
    for (const auto& sharedString : sharedStrings)
        sharedStringsMap[sharedString] = currentSharedStringIndex++;
    return sharedStringsMap;
}
