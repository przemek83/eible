#include "ExportDsvTest.h"

#include <ExportDsv.h>
#include <QBuffer>
#include <QTableView>
#include <QTest>

#include "TestTableModel.h"

QString ExportDsvTest::tableSheetData_ = R"()";
QString ExportDsvTest::multiSelectionTableSheetData_ = R"()";
QString ExportDsvTest::headersOnlySheetData_ = R"()";
QString ExportDsvTest::emptySheetData_ = "";
QStringList ExportDsvTest::headers_{"Text", "Numeric", "Date"};

void ExportDsvTest::initTestCase() {}

void ExportDsvTest::testExportingEmptyTable()
{
    TestTableModel model(0, 0);
    QTableView view;
    view.setModel(&model);

    QByteArray exportedByteArray;
    QBuffer exportedBuffer(&exportedByteArray);
    exportedBuffer.open(QIODevice::WriteOnly);

    ExportDsv exportDsv('\t');
    exportDsv.exportView(view, exportedBuffer);

    QCOMPARE(exportedByteArray, emptySheetData_);
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
