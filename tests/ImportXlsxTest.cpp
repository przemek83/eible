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
                                           "2020-01-03",
                                           "Other item",
                                           "2020-01-04",
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
