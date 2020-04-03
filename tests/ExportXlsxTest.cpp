#include "ExportXlsxTest.h"

#include <ExportXlsx.h>
#include <QCryptographicHash>
#include <QSignalSpy>
#include <QTest>
#include <QTableWidget>
#include <quazip5/quazip.h>
#include <quazip5/quazipfile.h>

#include "Utilities.h"

namespace
{
std::string tableSheetData =
    R"(<sheetData>)"
    R"(<row r="1" spans="1:1" x14ac:dyDescent="0.25">)"
    R"(<c r="A1" t="str" s="6"><v>Text</v></c>)"
    R"(<c r="B1" t="str" s="6"><v>Numeric</v></c>)"
    R"(<c r="C1" t="str" s="6"><v>Date</v></c>)"
    R"(</row>)"
    R"(<row r="2" spans="1:1" x14ac:dyDescent="0.25">)"
    R"(<c r="A2" t="str"><v>Item 0,0</v></c>)"
    R"(<c r="B2" s="4"><v>10</v></c>)"
    R"(<c r="C2" s="3"><v>43891</v></c>)"
    R"(</row>)"
    R"(<row r="3" spans="1:1" x14ac:dyDescent="0.25">)"
    R"(<c r="A3" t="str"><v>Item 0,1</v></c>)"
    R"(<c r="B3" s="4"><v>11</v></c>)"
    R"(<c r="C3" s="3"><v>43892</v></c>)"
    R"(</row>)"
    R"(<row r="4" spans="1:1" x14ac:dyDescent="0.25">)"
    R"(<c r="A4" t="str"><v>Item 0,2</v></c>)"
    R"(<c r="B4" s="4"><v>12</v></c>)"
    R"(<c r="C4" s="3"><v>43893</v></c>)"
    R"(</row>)"
    R"(</sheetData>)";

std::string headersOnlySheetData =
    R"(<sheetData>)"
    R"(<row r="1" spans="1:1" x14ac:dyDescent="0.25">)"
    R"(<c r="A1" t="str" s="6"><v>Text</v></c>)"
    R"(<c r="B1" t="str" s="6"><v>Numeric</v></c>)"
    R"(<c r="C1" t="str" s="6"><v>Date</v></c>)"
    R"(</row>)"
    R"(</sheetData>)";

std::string emptySheetData = R"(</sheetData>)";

QStringList headers {"Text", "Numeric", "Date"};
}

QByteArray ExportXlsxTest::retrieveFileFromZip(const QString& zipFilePath,
                                               const QString& fileName) const
{
    QuaZip inZip(zipFilePath);
    inZip.open(QuaZip::mdUnzip);
    inZip.setCurrentFile(fileName);
    QuaZipFile inZipFile(&inZip);
    inZipFile.open(QIODevice::ReadOnly | QIODevice::Text);
    return inZipFile.readAll();
}

void ExportXlsxTest::initTable(QTableWidget& tableWidget) const
{
    const int columnCount {3};
    const int rowCount {3};
    tableWidget.setRowCount(rowCount);
    tableWidget.setColumnCount(columnCount);
    tableWidget.setHorizontalHeaderLabels(headers);
    for (int column = 0; column < columnCount; ++column)
        for (int row = 0; row < rowCount; ++row)
        {
            auto item = new QTableWidgetItem("Item " + QString::number(column) +
                                             "," + QString::number(row));
            switch (column)
            {
                case 0:
                    item->setData(Qt::DisplayRole, item->text());
                    break;
                case 1:
                    item->setData(Qt::DisplayRole, 10 * column + row);
                    break;
                case 2:
                    item->setData(Qt::DisplayRole,
                                  QDate(2020, 1 + column, 1 + row));
                    break;
            }
            tableWidget.setItem(row, column, item);
        }
}

void ExportXlsxTest::testExportingEmptyTable()
{
    QTableWidget tableWidget;
    const QString testFilePath(QCoreApplication::applicationDirPath() +
                               "/test1.xlsx");
    ExportXlsx exportXlsx(testFilePath);
    exportXlsx.exportView(&tableWidget);

    const QByteArray actual =
        retrieveFileFromZip(testFilePath, "xl/worksheets/sheet1.xml");
    const QByteArray expected =
        Utilities::composeXlsxSheet(QString::fromStdString(emptySheetData));

    QCOMPARE(actual, expected);
}

void ExportXlsxTest::testExportingHeadersOnly()
{
    QTableWidget tableWidget;
    const QString testFilePath(QCoreApplication::applicationDirPath() +
                               "/test1.xlsx");
    const int columnCount {3};
    const int rowCount {0};
    tableWidget.setRowCount(rowCount);
    tableWidget.setColumnCount(columnCount);
    tableWidget.setHorizontalHeaderLabels(headers);

    ExportXlsx exportXlsx(testFilePath);
    exportXlsx.exportView(&tableWidget);

    const QByteArray actual =
        retrieveFileFromZip(testFilePath, "xl/worksheets/sheet1.xml");
    const QByteArray expected =
        Utilities::composeXlsxSheet(QString::fromStdString(headersOnlySheetData));

    QCOMPARE(actual, expected);
}

void ExportXlsxTest::testExportingSimpleTable()
{
    QTableWidget tableWidget;
    initTable(tableWidget);

    const QString testFilePath(QCoreApplication::applicationDirPath() +
                               "/test1.xlsx");
    ExportXlsx exportXlsx(testFilePath);
    exportXlsx.exportView(&tableWidget);

    const QByteArray actual =
        retrieveFileFromZip(testFilePath, "xl/worksheets/sheet1.xml");
    const QByteArray expected =
        Utilities::composeXlsxSheet(QString::fromStdString(tableSheetData));

    QCOMPARE(actual, expected);
}
