#include "ExportXlsxTest.h"

#include <ExportXlsx.h>
#include <QCryptographicHash>
#include <QSignalSpy>
#include <QTableView>
#include <QTest>
#include <quazip5/quazip.h>
#include <quazip5/quazipfile.h>

#include "TestTableModel.h"
#include "Utilities.h"

QString ExportXlsxTest::zipWorkSheetPath_ {"xl/worksheets/sheet1.xml"};

QString ExportXlsxTest::tableSheetData_ =
    R"(<sheetData>)"
    R"(<row r="1" spans="1:1" x14ac:dyDescent="0.25">)"
    R"(<c r="A1" t="str" s="6"><v>Text</v></c>)"
    R"(<c r="B1" t="str" s="6"><v>Numeric</v></c>)"
    R"(<c r="C1" t="str" s="6"><v>Date</v></c>)"
    R"(</row>)"
    R"(<row r="2" spans="1:1" x14ac:dyDescent="0.25">)"
    R"(<c r="A2" t="str"><v>Item 0, 0</v></c>)"
    R"(<c r="B2" s="4"><v>1</v></c>)"
    R"(<c r="C2" s="3"><v>43833</v></c>)"
    R"(</row>)"
    R"(<row r="3" spans="1:1" x14ac:dyDescent="0.25">)"
    R"(<c r="A3" t="str"><v>Item 0, 1</v></c>)"
    R"(<c r="B3" s="4"><v>2</v></c>)"
    R"(<c r="C3" s="3"><v>43834</v></c>)"
    R"(</row>)"
    R"(<row r="4" spans="1:1" x14ac:dyDescent="0.25">)"
    R"(<c r="A4" t="str"><v>Item 0, 2</v></c>)"
    R"(<c r="B4" s="4"><v>3</v></c>)"
    R"(<c r="C4" s="3"><v>43835</v></c>)"
    R"(</row>)"
    R"(</sheetData>)";

QString ExportXlsxTest::multiSelectionTableSheetData_ =
    R"(<sheetData>)"
    R"(<row r="1" spans="1:1" x14ac:dyDescent="0.25">)"
    R"(<c r="A1" t="str" s="6"><v>Text</v></c>)"
    R"(<c r="B1" t="str" s="6"><v>Numeric</v></c>)"
    R"(<c r="C1" t="str" s="6"><v>Date</v></c>)"
    R"(</row>)"
    R"(<row r="2" spans="1:1" x14ac:dyDescent="0.25">)"
    R"(<c r="A2" t="str"><v>Item 0, 0</v></c>)"
    R"(<c r="B2" s="4"><v>1</v></c>)"
    R"(<c r="C2" s="3"><v>43833</v></c>)"
    R"(</row>)"
    R"(<row r="3" spans="1:1" x14ac:dyDescent="0.25">)"
    R"(<c r="A3" t="str"><v>Item 0, 2</v></c>)"
    R"(<c r="B3" s="4"><v>3</v></c>)"
    R"(<c r="C3" s="3"><v>43835</v></c>)"
    R"(</row>)"
    R"(</sheetData>)";

QString ExportXlsxTest::headersOnlySheetData_ =
    R"(<sheetData>)"
    R"(<row r="1" spans="1:1" x14ac:dyDescent="0.25">)"
    R"(<c r="A1" t="str" s="6"><v>Text</v></c>)"
    R"(<c r="B1" t="str" s="6"><v>Numeric</v></c>)"
    R"(<c r="C1" t="str" s="6"><v>Date</v></c>)"
    R"(</row>)"
    R"(</sheetData>)";

QString ExportXlsxTest::emptySheetData_ = R"(</sheetData>)";

QStringList ExportXlsxTest::headers_ {"Text", "Numeric", "Date"};

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

void ExportXlsxTest::compareWorkSheets(const QString& testFilePath,
                                       const QString& sheetData) const
{
    const QByteArray actual = retrieveFileFromZip(testFilePath, zipWorkSheetPath_);
    const QByteArray expected = Utilities::composeXlsxSheet(sheetData);
    QCOMPARE(actual, expected);
}

void ExportXlsxTest::exportZip(const QAbstractItemView* view,
                               const QString& testFilePath) const
{
    ExportXlsx exportXlsx(testFilePath);
    exportXlsx.exportView(view);
}

QString ExportXlsxTest::getTestFilePath() const
{
    return QCoreApplication::applicationDirPath() + "/test1.xlsx";
}

void ExportXlsxTest::initTestCase()
{
}

void ExportXlsxTest::testExportingEmptyTable()
{
    TestTableModel model(0, 0);
    QTableView view;
    view.setModel(&model);
    const QString testFilePath {getTestFilePath()};
    exportZip(&view, testFilePath);
    compareWorkSheets(testFilePath, emptySheetData_);
}

void ExportXlsxTest::testExportingHeadersOnly()
{
    TestTableModel model(3, 0);
    QTableView view;
    view.setModel(&model);
    const QString testFilePath {getTestFilePath()};
    exportZip(&view, testFilePath);
    compareWorkSheets(testFilePath, headersOnlySheetData_);
}

void ExportXlsxTest::testExportingSimpleTable()
{
    TestTableModel model(3, 3);
    QTableView view;
    view.setModel(&model);
    const QString testFilePath {getTestFilePath()};
    exportZip(&view, testFilePath);
    compareWorkSheets(testFilePath, tableSheetData_);
}

void ExportXlsxTest::testExportingViewWithMultiSelection()
{
    TestTableModel model(3, 3);
    QTableView view;
    view.setModel(&model);
    view.setSelectionBehavior(QAbstractItemView::SelectRows);
    view.setSelectionMode(QAbstractItemView::MultiSelection);
    view.selectionModel()->select(model.index(0, 0), QItemSelectionModel::Select);
    view.selectionModel()->select(model.index(2, 0), QItemSelectionModel::Select);
    const QString testFilePath {getTestFilePath()};
    exportZip(&view, testFilePath);
    compareWorkSheets(testFilePath, multiSelectionTableSheetData_);
}

void ExportXlsxTest::Benchmark_data()
{
    tableModelForBenchmarking_ = new TestTableModel(100, 100000);
}

void ExportXlsxTest::Benchmark()
{
    QTableView view;
    view.setModel(tableModelForBenchmarking_);
    const QString testFilePath {getTestFilePath()};
    QBENCHMARK
    {
        exportZip(&view, testFilePath);
    }
}

void ExportXlsxTest::cleanupTestCase()
{
    delete tableModelForBenchmarking_;
}
