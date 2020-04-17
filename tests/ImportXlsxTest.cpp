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

QSet<QString> ImportXlsxTest::sharedStrings_{"Trait #1",
                                             "n",
                                             "Units",
                                             "mass (kg)",
                                             "pow",
                                             "trait 8",
                                             "trait 22",
                                             "trait 5",
                                             "Transaction date",
                                             "trait 16",
                                             "trait 6",
                                             "b",
                                             "cena_m",
                                             "Pow",
                                             "Price",
                                             "Q25",
                                             "Sxx",
                                             "date",
                                             "Sxy",
                                             "Mediana:",
                                             "cena metra",
                                             "text",
                                             "trait 17",
                                             "brown",
                                             "Syy",
                                             "trait 19",
                                             "Q90",
                                             "yellow",
                                             "trait 21",
                                             "pink",
                                             "Price per unit",
                                             "trait 1",
                                             "white",
                                             "a (beta w wiki)",
                                             "cena nier",
                                             "y",
                                             "Średnia:",
                                             "black",
                                             "trait 2",
                                             "2020-01-04",
                                             "Date",
                                             "name",
                                             "modificator",
                                             "Other item",
                                             "Q75",
                                             "c",
                                             "trait 4",
                                             "y=ax+b",
                                             "blue",
                                             "x",
                                             "Value #1",
                                             "b (alfa w wiki)",
                                             "Numeric",
                                             "trait 3",
                                             "Item 0, 0",
                                             "trait 20",
                                             "Text",
                                             "Sx",
                                             "trait 7",
                                             "trait 18",
                                             "a",
                                             "Q50",
                                             "data transakcji",
                                             "Sy",
                                             "Q10",
                                             "Another trait",
                                             "2020-01-03",
                                             "height",
                                             "red",
                                             "Cena"};

void ImportXlsxTest::testRetrievingSheetNames()
{
    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx ImportXlsx(xlsxTestFile);
    auto [success, actualNames] = ImportXlsx.getSheetList();
    QCOMPARE(success, true);
    QCOMPARE(actualNames, sheetNames_);
}

void ImportXlsxTest::testRetrievingSheetNamesFromEmptyFile()
{
    QByteArray byteArray;
    QBuffer emptyBuffer(&byteArray);
    ImportXlsx ImportXlsx(emptyBuffer);
    auto [success, actualNames] = ImportXlsx.getSheetList();
    QCOMPARE(success, false);
    QCOMPARE(actualNames, {});
}

void ImportXlsxTest::testGetStyles()
{
    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx ImportXlsx(xlsxTestFile);
    auto [success, dateStyle, allStyles] = ImportXlsx.getStyles();
    QCOMPARE(success, true);
    QCOMPARE(dateStyle, QList({14, 15, 16, 17, 22, 167, 169, 170, 171}));
    QCOMPARE(allStyles, QList({164, 165, 164, 166, 167, 168, 169, 164, 164, 170,
                               171, 164}));
}

void ImportXlsxTest::testGetStylesNoContent()
{
    QFile xlsxTestFile(QStringLiteral(":/template.xlsx"));
    ImportXlsx ImportXlsx(xlsxTestFile);
    auto [success, dateStyle, allStyles] = ImportXlsx.getStyles();
    QCOMPARE(success, true);
    QCOMPARE(dateStyle, QList({14, 15, 16, 17, 22}));
    QCOMPARE(allStyles, QList({0, 164, 10, 14, 4, 0, 0, 3}));
}

void ImportXlsxTest::testGetSharedStrings()
{
    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx ImportXlsx(xlsxTestFile);
    auto [success, actualSharedStrings] = ImportXlsx.getSharedStrings();
    QCOMPARE(success, true);
    QCOMPARE(actualSharedStrings, sharedStrings_);
}

void ImportXlsxTest::testGetSharedStringsNoContent()
{
    QFile xlsxTestFile(QStringLiteral(":/template.xlsx"));
    ImportXlsx ImportXlsx(xlsxTestFile);
    auto [success, actualSharedStrings] = ImportXlsx.getSharedStrings();
    QCOMPARE(success, true);
    QCOMPARE(actualSharedStrings, {});
}
