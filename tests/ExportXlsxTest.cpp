#include "ExportXlsxTest.h"

#include <ExportXlsx.h>
#include <quazip/quazip.h>
#include <quazip/quazipfile.h>

#include <QBuffer>
#include <QCryptographicHash>
#include <QTableView>
#include <QTest>

#include "TestTableModel.h"
#include "Utilities.h"

ExportXlsxTest::ExportXlsxTest(QObject* parent) : QObject(parent) {}

QByteArray ExportXlsxTest::retrieveFileFromZip(QBuffer& exportedZip,
                                               const QString& fileName)
{
    QuaZip inZip(&exportedZip);
    inZip.open(QuaZip::mdUnzip);
    inZip.setCurrentFile(fileName);
    QuaZipFile inZipFile(&inZip);
    inZipFile.open(QIODevice::ReadOnly | QIODevice::Text);
    return inZipFile.readAll();
}

void ExportXlsxTest::compareWorkSheets(QBuffer& exportedZip,
                                       const QString& sheetData) const
{
    const QByteArray actual =
        retrieveFileFromZip(exportedZip, zipWorkSheetPath_);
    const QByteArray expected = Utilities::composeXlsxSheet(sheetData);
    QCOMPARE(actual, expected);
}

void ExportXlsxTest::exportZip(const QAbstractItemView& view,
                               QBuffer& exportedZip)
{
    exportedZip.open(QIODevice::WriteOnly);
    ExportXlsx exportXlsx;
    exportXlsx.exportView(view, exportedZip);
}

void ExportXlsxTest::initTestCase() {}

void ExportXlsxTest::testExportingEmptyTable()
{
    TestTableModel model(0, 0);
    QTableView view;
    view.setModel(&model);

    QByteArray exportedZipBuffer;
    QBuffer exportedZip(&exportedZipBuffer);
    exportZip(view, exportedZip);

    compareWorkSheets(exportedZip, emptySheetData_);
}

void ExportXlsxTest::testExportingHeadersOnly()
{
    TestTableModel model(3, 0);
    QTableView view;
    view.setModel(&model);

    QByteArray exportedZipBuffer;
    QBuffer exportedZip(&exportedZipBuffer);
    exportZip(view, exportedZip);

    compareWorkSheets(exportedZip, headersOnlySheetData_);
}

void ExportXlsxTest::testExportingSimpleTable()
{
    TestTableModel model(3, 3);
    QTableView view;
    view.setModel(&model);

    QByteArray exportedZipBuffer;
    QBuffer exportedZip(&exportedZipBuffer);
    exportZip(view, exportedZip);

    compareWorkSheets(exportedZip, tableSheetData_);
}

void ExportXlsxTest::testExportingViewWithMultiSelection()
{
    TestTableModel model(3, 3);
    QTableView view;
    view.setModel(&model);
    view.setSelectionBehavior(QAbstractItemView::SelectRows);
    view.setSelectionMode(QAbstractItemView::MultiSelection);
    view.selectionModel()->select(model.index(0, 0),
                                  QItemSelectionModel::Select);
    view.selectionModel()->select(model.index(2, 0),
                                  QItemSelectionModel::Select);

    QByteArray exportedZipBuffer;
    QBuffer exportedZip(&exportedZipBuffer);
    exportZip(view, exportedZip);

    compareWorkSheets(exportedZip, multiSelectionTableSheetData_);
}

void ExportXlsxTest::benchmark_data()
{
    QSKIP("Skip preparing benchmark data.");
    tableModelForBenchmarking_ = new TestTableModel(100, 100000);
}

void ExportXlsxTest::benchmark()
{
    QSKIP("Skip benchmark.");
    QTableView view;
    view.setModel(tableModelForBenchmarking_);
    QBENCHMARK
    {
        QByteArray exportedZipBuffer;
        QBuffer buffer(&exportedZipBuffer);
        exportZip(view, buffer);
    }
}

void ExportXlsxTest::cleanupTestCase() { delete tableModelForBenchmarking_; }
