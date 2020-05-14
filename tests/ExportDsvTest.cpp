#include "ExportDsvTest.h"

#include <ExportDsv.h>
#include <QBuffer>
#include <QTableView>
#include <QTest>

#include "TestTableModel.h"

ExportDsvTest::ExportDsvTest(QObject* parent) : QObject(parent) {}

void ExportDsvTest::testEmptyTable_data()
{
    QTest::addColumn<char>("separator");

    QTest::newRow("TSV empty table") << '\t';
    QTest::newRow("CSV empty table") << ',';
}

void ExportDsvTest::testEmptyTable()
{
    QFETCH(char, separator);

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

void ExportDsvTest::testHeadersOnly_data()
{
    QTest::addColumn<char>("separator");
    QTest::addColumn<QString>("expected");

    QTest::newRow("TSV headers only") << '\t' << headersOnlyDataTsv_;
    QTest::newRow("CSV headers only") << ',' << headersOnlyDataCsv_;
}

void ExportDsvTest::testHeadersOnly()
{
    QFETCH(char, separator);
    QFETCH(QString, expected);

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

void ExportDsvTest::testSimpleTable_data()
{
    QTest::addColumn<char>("separator");
    QTest::addColumn<QString>("expected");

    QTest::newRow("TSV simple table") << '\t' << tableDataTsv_;
    QTest::newRow("CSV simple table") << ',' << tableDataCsv_;
}

void ExportDsvTest::testSimpleTable()
{
    QFETCH(char, separator);
    QFETCH(QString, expected);

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

void ExportDsvTest::testViewWithMultiSelection_data()
{
    QTest::addColumn<char>("separator");
    QTest::addColumn<QString>("expected");

    QTest::newRow("TSV multi selection") << '\t' << multiSelectionTableDataTsv_;
    QTest::newRow("CSV multi selection") << ',' << multiSelectionTableDataCsv_;
}
void ExportDsvTest::testViewWithMultiSelection()
{
    QFETCH(char, separator);
    QFETCH(QString, expected);

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

void ExportDsvTest::testSpecialCharInStringField_data()
{
    QTest::addColumn<char>("separator");
    QTest::addColumn<char>("specialChar");
    QTest::addColumn<QString>("expected");

    QTest::newRow("TSV separator in string")
        << '\t' << '\t' << separatorInStringFieldDataTsv_;
    QTest::newRow("CSV separator in string")
        << ',' << ',' << separatorInStringFieldDataCsv_;

    QTest::newRow("TSV new line in string") << '\t' << '\n'
                                            << newLineInStringFieldDataTsv_;
    QTest::newRow("CSV new line in string") << ',' << '\n'
                                            << newLineInStringFieldDataCsv_;

    QTest::newRow("TSV double quotes in string")
        << '\t' << '\"' << doubleQuotesInStringFieldDataTsv_;
    QTest::newRow("CSV double quotes in string")
        << ',' << '\"' << doubleQuotesInStringFieldDataCsv_;
}

void ExportDsvTest::testSpecialCharInStringField()
{
    QFETCH(char, separator);
    QFETCH(char, specialChar);
    QFETCH(QString, expected);

    TestTableModel model(3, 2);
    QTableView view;
    view.setModel(&model);
    model.setData(model.index(0, 0), "Other" + QString(specialChar) + "item");

    QByteArray exportedByteArray;
    QBuffer exportedBuffer(&exportedByteArray);
    exportedBuffer.open(QIODevice::WriteOnly);

    ExportDsv exportDsv(separator);
    exportDsv.exportView(view, exportedBuffer);

    QCOMPARE(exportedByteArray, expected);
}

void ExportDsvTest::testCustomDateFormat()
{
    TestTableModel model(3, 1);
    QTableView view;
    view.setModel(&model);

    QByteArray exportedByteArray;
    QBuffer exportedBuffer(&exportedByteArray);
    exportedBuffer.open(QIODevice::WriteOnly);

    ExportDsv exportDsv(',');
    exportDsv.setDateFormat(QStringLiteral("dd/MM/yy"));
    exportDsv.exportView(view, exportedBuffer);

    QCOMPARE(exportedByteArray, customDateFormatData_);
}

void ExportDsvTest::testDefaultLocaleShortDate()
{
    TestTableModel model(3, 1);
    QTableView view;
    view.setModel(&model);

    QByteArray exportedByteArray;
    QBuffer exportedBuffer(&exportedByteArray);
    exportedBuffer.open(QIODevice::WriteOnly);

    ExportDsv exportDsv(',');
    exportDsv.setDateFormat(Qt::DefaultLocaleShortDate);
    exportDsv.exportView(view, exportedBuffer);

    QCOMPARE(exportedByteArray, defaultLocaleShortDateData_);
}

void ExportDsvTest::testLocaleForNumbers()
{
    TestTableModel model(3, 1);
    QTableView view;
    view.setModel(&model);

    QByteArray exportedByteArray;
    QBuffer exportedBuffer(&exportedByteArray);
    exportedBuffer.open(QIODevice::WriteOnly);

    ExportDsv exportDsv('\t');
    QLocale locale(QLocale::Polish, QLocale::Poland);
    exportDsv.setNumbersLocale(locale);
    exportDsv.exportView(view, exportedBuffer);

    QCOMPARE(exportedByteArray, localeForNumbersData_);
}

void ExportDsvTest::benchmark_data()
{
    QSKIP("Skip preparing benchmark data.");
    tableModelForBenchmarking_ = new TestTableModel(100, 100000);
}

void ExportDsvTest::benchmark()
{
    QSKIP("Skip benchmark.");
    QTableView view;
    view.setModel(tableModelForBenchmarking_);
    QBENCHMARK
    {
        QByteArray exportedByteArray;
        QBuffer exportedBuffer(&exportedByteArray);
        exportedBuffer.open(QIODevice::WriteOnly);

        ExportDsv exportDsv(',');
        exportDsv.exportView(view, exportedBuffer);
    }
}

void ExportDsvTest::cleanupTestCase() { delete tableModelForBenchmarking_; }
