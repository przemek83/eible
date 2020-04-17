#include "ImportXlsxTest.h"

#include <QBuffer>
#include <QTest>

#include "ImportXlsx.h"

QMap<QString, QString> ImportXlsxTest::sheetNames_ = {
    {"Sheet1", "xl/worksheets/sheet1.xml"},
    {"Sheet2", "xl/worksheets/sheet2.xml"},
    {"Sheet3(empty)", "xl/worksheets/sheet3.xml"},
    {"Sheet4", "xl/worksheets/sheet4.xml"},
    {"Sheet5", "xl/worksheets/sheet5.xml"},
    {"Sheet6", "xl/worksheets/sheet6.xml"},
    {"Sheet7", "xl/worksheets/sheet7.xml"}};

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
