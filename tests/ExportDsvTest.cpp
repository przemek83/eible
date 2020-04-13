#include "ExportDsvTest.h"

#include <ExportDsv.h>
#include <QBuffer>
#include <QTableView>
#include <QTest>

#include "TestTableModel.h"

QString ExportDsvTest::tableDataTsv_ =
    "Text\tNumeric\tDate\nItem 0, 0\t1.00\t2020-01-03\nItem 0, "
    "1\t2.00\t2020-01-04\nItem 0, 2\t3.00\t2020-01-05\n";
QString ExportDsvTest::tableDataCsv_ =
    "Text;Numeric;Date\n\"Item 0, 0\";1.00;2020-01-03\n\"Item 0, "
    "1\";2.00;2020-01-04\n\"Item 0, 2\";3.00;2020-01-05\n";
QString ExportDsvTest::multiSelectionTableDataTsv_ =
    "Text\tNumeric\tDate\nItem 0, 0\t1.00\t2020-01-03\nItem 0, "
    "2\t3.00\t2020-01-05\n";
QString ExportDsvTest::multiSelectionTableDataCsv_ =
    "Text;Numeric;Date\n\"Item 0, 0\";1.00;2020-01-03\n\"Item 0, "
    "2\";3.00;2020-01-05\n";
QString ExportDsvTest::headersOnlyDataTsv_ = "Text\tNumeric\tDate\n";
QString ExportDsvTest::headersOnlyDataCsv_ = "Text;Numeric;Date\n";
QString ExportDsvTest::emptyData_ = "";
QStringList ExportDsvTest::headers_{"Text", "Numeric", "Date"};

void ExportDsvTest::initTestCase() {}

void ExportDsvTest::checkExportingEmptyTable(char separator)
{
    TestTableModel model(0, 0);
    QTableView view;
    view.setModel(&model);

    QByteArray exportedByteArray;
    QBuffer exportedBuffer(&exportedByteArray);
    exportedBuffer.open(QIODevice::WriteOnly);

    ExportDsv exportDsv(separator);
    exportDsv.exportView(view, exportedBuffer);

    QCOMPARE(exportedByteArray, emptyData_);
}

void ExportDsvTest::testExportingEmptyTableTsv()
{
    checkExportingEmptyTable('\t');
}

void ExportDsvTest::testExportingEmptyTableCsv()
{
    checkExportingEmptyTable(';');
}

void ExportDsvTest::checkExportingHeadersOnly(char separator,
                                              const QString& expected)
{
    TestTableModel model(3, 0);
    QTableView view;
    view.setModel(&model);

    QByteArray exportedByteArray;
    QBuffer exportedBuffer(&exportedByteArray);
    exportedBuffer.open(QIODevice::WriteOnly);

    ExportDsv exportDsv(separator);
    exportDsv.exportView(view, exportedBuffer);

    QCOMPARE(exportedByteArray, expected);
}

void ExportDsvTest::testExportingHeadersOnlyTsv()
{
    checkExportingHeadersOnly('\t', headersOnlyDataTsv_);
}

void ExportDsvTest::testExportingHeadersOnlyCsv()
{
    checkExportingHeadersOnly(';', headersOnlyDataCsv_);
}

void ExportDsvTest::checkExportingSimpleTable(char separator,
                                              const QString& expected)
{
    TestTableModel model(3, 3);
    QTableView view;
    view.setModel(&model);

    QByteArray exportedByteArray;
    QBuffer exportedBuffer(&exportedByteArray);
    exportedBuffer.open(QIODevice::WriteOnly);

    ExportDsv exportDsv(separator);
    exportDsv.exportView(view, exportedBuffer);

    QCOMPARE(exportedByteArray, expected);
}

void ExportDsvTest::testExportingSimpleTableTsv()
{
    checkExportingSimpleTable('\t', tableDataTsv_);
}

void ExportDsvTest::testExportingSimpleTableCsv()
{
    checkExportingSimpleTable(';', tableDataCsv_);
}

void ExportDsvTest::checkExportingViewWithMultiSelection(
    char separator, const QString& expected)
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

    QByteArray exportedByteArray;
    QBuffer exportedBuffer(&exportedByteArray);
    exportedBuffer.open(QIODevice::WriteOnly);

    ExportDsv exportDsv(separator);
    exportDsv.exportView(view, exportedBuffer);

    QCOMPARE(exportedByteArray, expected);
}

void ExportDsvTest::testExportingViewWithMultiSelectionTsv()
{
    checkExportingViewWithMultiSelection('\t', multiSelectionTableDataTsv_);
}

void ExportDsvTest::testExportingViewWithMultiSelectionCsv()
{
    checkExportingViewWithMultiSelection(';', multiSelectionTableDataCsv_);
}

void ExportDsvTest::Benchmark_data()
{
    tableModelForBenchmarking_ = new TestTableModel(100, 100000);
}

void ExportDsvTest::Benchmark()
{
    //    QSKIP("Skip benchmark.");
    QTableView view;
    view.setModel(tableModelForBenchmarking_);
    QBENCHMARK
    {
        QByteArray exportedByteArray;
        QBuffer exportedBuffer(&exportedByteArray);
        exportedBuffer.open(QIODevice::WriteOnly);

        ExportDsv exportDsv(';');
        exportDsv.exportView(view, exportedBuffer);
    }
}

void ExportDsvTest::cleanupTestCase() { delete tableModelForBenchmarking_; }
