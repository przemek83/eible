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

QList<QStringList> ImportXlsxTest::testColumnNames_ = {
    {"Text", "Numeric", "Date"},
    {"Trait #1", "Value #1", "Transaction date", "Units", "Price",
     "Price per unit", "Another trait"},
    {},
    {"cena nier", "pow", "cena metra", "data transakcji", "text"},
    {"name", "date", "mass (kg)", "height", "---", "---", "---", "---", "---",
     "---", "---", "---"},
    {"modificator", "x", "y"},
    {"Pow", "Cena", "cena_m", "---", "---", "---"}};

std::vector<unsigned int> ImportXlsxTest::expectedRowCounts_{2,  19,  0, 4,
                                                             30, 100, 71};

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
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<QStringList>("expectedColumnList");
    for (int i = 0; i < testColumnNames_.size(); ++i)
    {
        const QString& sheetName{sheets_[i].first};
        QTest::newRow(("Columns in " + sheetName).toStdString().c_str())
            << sheetName << testColumnNames_[i];
    }
}

void ImportXlsxTest::testGetColumnList()
{
    QFETCH(QString, sheetName);
    QFETCH(QStringList, expectedColumnList);
    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    importXlsx.setSharedStrings(sharedStrings_);
    importXlsx.setSheets(sheets_);
    auto [success, actualColumnList] = importXlsx.getColumnNames(sheetName);
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
    auto [success, actualColumnList] =
        importXlsx.getColumnNames(sheets_[4].first);

    std::list<QString> expectedColumnList(testColumnNames_[4].size());
    std::replace_copy(testColumnNames_[4].begin(), testColumnNames_[4].end(),
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
    auto [success, actualColumnList] =
        importXlsx.getColumnNames(sheets_[4].first);
    QCOMPARE(success, true);
    QCOMPARE(actualColumnList, testColumnNames_[4]);

    std::tie(success, actualColumnList) =
        importXlsx.getColumnNames(sheets_[0].first);
    QCOMPARE(success, true);
    QCOMPARE(actualColumnList, testColumnNames_[0]);
}

void ImportXlsxTest::testGetColumnTypes_data()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<QVector<ColumnType>>("expectedColumnTypes");

    QTest::newRow(("Column types in " + sheets_[0].first).toStdString().c_str())
        << sheets_[0].first
        << QVector({ColumnType::STRING, ColumnType::NUMBER, ColumnType::DATE});
    QTest::newRow(("Column types in " + sheets_[1].first).toStdString().c_str())
        << sheets_[1].first
        << QVector({ColumnType::STRING, ColumnType::NUMBER, ColumnType::DATE,
                    ColumnType::NUMBER, ColumnType::NUMBER, ColumnType::NUMBER,
                    ColumnType::STRING});
    QTest::newRow(("Column types in " + sheets_[2].first).toStdString().c_str())
        << sheets_[2].first << QVector<ColumnType>();
    QTest::newRow(("Column types in " + sheets_[3].first).toStdString().c_str())
        << sheets_[3].first
        << QVector({ColumnType::NUMBER, ColumnType::NUMBER, ColumnType::NUMBER,
                    ColumnType::DATE, ColumnType::STRING});
    QTest::newRow(("Column types in " + sheets_[4].first).toStdString().c_str())
        << sheets_[4].first
        << QVector({ColumnType::STRING, ColumnType::STRING, ColumnType::NUMBER,
                    ColumnType::NUMBER, ColumnType::STRING, ColumnType::NUMBER,
                    ColumnType::STRING, ColumnType::STRING, ColumnType::NUMBER,
                    ColumnType::NUMBER, ColumnType::NUMBER,
                    ColumnType::NUMBER});
    QTest::newRow(("Column types in " + sheets_[5].first).toStdString().c_str())
        << sheets_[5].first
        << QVector(
               {ColumnType::NUMBER, ColumnType::NUMBER, ColumnType::NUMBER});
    QTest::newRow(("Column types in " + sheets_[6].first).toStdString().c_str())
        << sheets_[6].first
        << QVector({ColumnType::NUMBER, ColumnType::NUMBER, ColumnType::NUMBER,
                    ColumnType::STRING, ColumnType::STRING,
                    ColumnType::NUMBER});
}

void ImportXlsxTest::testGetColumnTypes()
{
    QFETCH(QString, sheetName);
    QFETCH(QVector<ColumnType>, expectedColumnTypes);

    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    importXlsx.setSharedStrings(sharedStrings_);
    importXlsx.setDateStyles(dateStyles_);
    importXlsx.setAllStyles(allStyles_);
    importXlsx.setSheets(sheets_);

    auto [success, columnTypes] = importXlsx.getColumnTypes(sheetName);
    QCOMPARE(success, true);
    QCOMPARE(columnTypes, expectedColumnTypes);
}

void ImportXlsxTest::testGetColumnCount_data()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<int>("expectedColumnCount");
    for (int i = 0; i < testColumnNames_.size(); ++i)
    {
        const QString& sheetName{sheets_[i].first};
        QTest::newRow(("Columns in " + sheetName).toStdString().c_str())
            << sheetName << testColumnNames_[i].size();
    }
}

void ImportXlsxTest::testGetColumnCount()
{
    QFETCH(QString, sheetName);
    QFETCH(int, expectedColumnCount);

    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    importXlsx.setSharedStrings(sharedStrings_);
    importXlsx.setDateStyles(dateStyles_);
    importXlsx.setAllStyles(allStyles_);
    importXlsx.setSheets(sheets_);
    auto [success, actualColumnCount] = importXlsx.getColumnCount(sheetName);
    QCOMPARE(success, true);
    QCOMPARE(actualColumnCount, expectedColumnCount);
}

void ImportXlsxTest::testGetRowCount_data()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<unsigned int>("expectedRowCount");
    for (int i = 0; i < testColumnNames_.size(); ++i)
    {
        const QString& sheetName{sheets_[i].first};
        QTest::newRow(("Rows in " + sheetName).toStdString().c_str())
            << sheetName << expectedRowCounts_[i];
    }
}

void ImportXlsxTest::testGetRowCount()
{
    QFETCH(QString, sheetName);
    QFETCH(unsigned int, expectedRowCount);

    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    importXlsx.setSharedStrings(sharedStrings_);
    importXlsx.setDateStyles(dateStyles_);
    importXlsx.setAllStyles(allStyles_);
    importXlsx.setSheets(sheets_);
    auto [success, actualRowCount] = importXlsx.getRowCount(sheetName);
    QCOMPARE(success, true);
    QCOMPARE(actualRowCount, expectedRowCount);
}

void ImportXlsxTest::testGetRowAndColumnCountViaGetColumnTypes_data()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<unsigned int>("expectedRowCount");
    QTest::addColumn<unsigned int>("expectedColumnCount");

    for (int i = 0; i < testColumnNames_.size(); ++i)
    {
        const QString& sheetName{sheets_[i].first};
        QTest::newRow(("Rows and columns via GetColumnTypes() in " + sheetName)
                          .toStdString()
                          .c_str())
            << sheetName << expectedRowCounts_[i]
            << static_cast<unsigned int>(testColumnNames_[i].size());
    }
}

void ImportXlsxTest::testGetRowAndColumnCountViaGetColumnTypes()
{
    QFETCH(QString, sheetName);
    QFETCH(unsigned int, expectedRowCount);
    QFETCH(unsigned int, expectedColumnCount);

    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    importXlsx.setSharedStrings(sharedStrings_);
    importXlsx.setDateStyles(dateStyles_);
    importXlsx.setAllStyles(allStyles_);
    importXlsx.setSheets(sheets_);
    importXlsx.getColumnTypes(sheetName);
    auto [successRowCount, actualRowCount] = importXlsx.getRowCount(sheetName);
    QCOMPARE(successRowCount, true);
    QCOMPARE(actualRowCount, expectedRowCount);
    auto [successColumnCount, actualColumnCount] =
        importXlsx.getColumnCount(sheetName);
    QCOMPARE(successColumnCount, true);
    QCOMPARE(actualColumnCount, expectedColumnCount);
}

void ImportXlsxTest::testGetData_data()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<QVector<QVector<QVariant>>>("expectedData");

    const QString& sheetName{sheets_[0].first};
    QTest::newRow(("Data in " + sheetName).toStdString().c_str())
        << sheetName
        << QVector<QVector<QVariant>>(
               {{3, 1., QDate(2020, 1, 3)}, {4, 2., QDate(2020, 1, 4)}});
}

void ImportXlsxTest::testGetData()
{
    QFETCH(QString, sheetName);
    QFETCH(QVector<QVector<QVariant>>, expectedData);

    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    importXlsx.setSharedStrings(sharedStrings_);
    importXlsx.setDateStyles(dateStyles_);
    importXlsx.setAllStyles(allStyles_);
    importXlsx.setSheets(sheets_);
    auto [success, actualData] = importXlsx.getData(sheetName, {});
    QCOMPARE(success, true);
    QCOMPARE(actualData, expectedData);
}
