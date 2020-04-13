#include "ExportDsvTest.h"

#include <ExportDsv.h>
#include <QBuffer>
#include <QTableView>
#include <QTest>

#include "TestTableModel.h"

QString ExportDsvTest::tableDataTsv_ =
    "Text\tNumeric\tDate\nItem 0, 0\t1.00\t2020-01-03\nItem 0, "
    "1\t2.00\t2020-01-04\nItem 0, 2\t3.00\t2020-01-05";
QString ExportDsvTest::tableDataCsv_ =
    "Text;Numeric;Date\nItem 0, 0;1.00;2020-01-03\nItem 0, "
    "1;2.00;2020-01-04\nItem 0, 2;3.00;2020-01-05";
QString ExportDsvTest::multiSelectionTableDataTsv_ =
    "Text\tNumeric\tDate\nItem 0, 0\t1.00\t2020-01-03\nItem 0, "
    "2\t3.00\t2020-01-05";
QString ExportDsvTest::multiSelectionTableDataCsv_ =
    "Text;Numeric;Date\nItem 0, 0;1.00;2020-01-03\nItem 0, "
    "2;3.00;2020-01-05";
QString ExportDsvTest::headersOnlyDataTsv_ = "Text\tNumeric\tDate";
QString ExportDsvTest::headersOnlyDataCsv_ = "Text;Numeric;Date";
QString ExportDsvTest::separatorInStringFieldDataTsv_ =
    "Text\tNumeric\tDate\n\"Other\titem\"\t1.00\t2020-01-03\nItem 0, "
    "1\t2.00\t2020-01-04";
QString ExportDsvTest::separatorInStringFieldDataCsv_ =
    "Text;Numeric;Date\n\"Other;item\";1.00;2020-01-03\nItem 0, "
    "1;2.00;2020-01-04";
QString ExportDsvTest::emptyData_ = "";
QStringList ExportDsvTest::headers_{"Text", "Numeric", "Date"};

void ExportDsvTest::initTestCase() {}

void ExportDsvTest::checkEmptyTable(char separator)
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

void ExportDsvTest::testEmptyTableTsv() { checkEmptyTable('\t'); }

void ExportDsvTest::testEmptyTableCsv() { checkEmptyTable(';'); }

void ExportDsvTest::checkHeadersOnly(char separator, const QString& expected)
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

void ExportDsvTest::testHeadersOnlyTsv()
{
    checkHeadersOnly('\t', headersOnlyDataTsv_);
}

void ExportDsvTest::testHeadersOnlyCsv()
{
    checkHeadersOnly(';', headersOnlyDataCsv_);
}

void ExportDsvTest::checkSimpleTable(char separator, const QString& expected)
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

void ExportDsvTest::testSimpleTableTsv()
{
    checkSimpleTable('\t', tableDataTsv_);
}

void ExportDsvTest::testSimpleTableCsv()
{
    checkSimpleTable(';', tableDataCsv_);
}

void ExportDsvTest::checkViewWithMultiSelection(char separator,
                                                const QString& expected)
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

void ExportDsvTest::testViewWithMultiSelectionTsv()
{
    checkViewWithMultiSelection('\t', multiSelectionTableDataTsv_);
}

void ExportDsvTest::testViewWithMultiSelectionCsv()
{
    checkViewWithMultiSelection(';', multiSelectionTableDataCsv_);
}

void ExportDsvTest::checkViewWithSeparatorInStringField(char separator,
                                                        const QString& expected)
{
    TestTableModel model(3, 2);
    QTableView view;
    view.setModel(&model);
    model.setData(model.index(0, 0), "Other" + QString(separator) + "item");

    QByteArray exportedByteArray;
    QBuffer exportedBuffer(&exportedByteArray);
    exportedBuffer.open(QIODevice::WriteOnly);

    ExportDsv exportDsv(separator);
    exportDsv.exportView(view, exportedBuffer);

    QCOMPARE(exportedByteArray, expected);
}

void ExportDsvTest::testViewWithSeparatorInStringFieldTsv()
{
    checkViewWithSeparatorInStringField('\t', separatorInStringFieldDataTsv_);
}

void ExportDsvTest::testViewWithSeparatorInStringFieldCsv()
{
    checkViewWithSeparatorInStringField(';', separatorInStringFieldDataCsv_);
}

void ExportDsvTest::Benchmark_data()
{
    tableModelForBenchmarking_ = new TestTableModel(100, 100000);
}

void ExportDsvTest::Benchmark()
{
    // QSKIP("Skip benchmark.");
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
