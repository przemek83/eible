#include "ExportDsvTest.h"

#include <ExportDsv.h>
#include <QBuffer>
#include <QTableView>
#include <QTest>

#include "TestTableModel.h"

QString ExportDsvTest::tableData_ = R"()";
QString ExportDsvTest::multiSelectionTableData_ = R"()";
QString ExportDsvTest::headersOnlyData_ = R"()";
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

void ExportDsvTest::testExportingHeadersOnly()
{
    TestTableModel model(3, 0);
    QTableView view;
    view.setModel(&model);

    //    QByteArray exportedZipBuffer;
    //    QBuffer exportedZip(&exportedZipBuffer);
    //    exportZip(view, exportedZip);

    //    compareWorkSheets(exportedZip, headersOnlySheetData_);
}

void ExportDsvTest::testExportingSimpleTable()
{
    TestTableModel model(3, 3);
    QTableView view;
    view.setModel(&model);

    //    QByteArray exportedZipBuffer;
    //    QBuffer exportedZip(&exportedZipBuffer);
    //    exportZip(view, exportedZip);

    //    compareWorkSheets(exportedZip, tableSheetData_);
}

void ExportDsvTest::testExportingViewWithMultiSelection()
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

    //    QByteArray exportedZipBuffer;
    //    QBuffer exportedZip(&exportedZipBuffer);
    //    exportZip(view, exportedZip);

    //    compareWorkSheets(exportedZip, multiSelectionTableSheetData_);
}
