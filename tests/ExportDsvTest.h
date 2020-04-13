#ifndef EXPORTDSVTEST_H
#define EXPORTDSVTEST_H

#include <QObject>

class ExportDsvTest : public QObject
{
    Q_OBJECT
private Q_SLOTS:
    void initTestCase();

    void testExportingEmptyTableTsv();
    void testExportingEmptyTableCsv();

    void testExportingHeadersOnly();

    void testExportingSimpleTable();

    void testExportingViewWithMultiSelection();

    //    void Benchmark_data();

    //    void Benchmark();

    //    void cleanupTestCase();

private:
    void checkExportingEmptyTable(char separator);

    static QString tableData_;
    static QString multiSelectionTableData_;
    static QString headersOnlyData_;
    static QString emptyData_;
    static QStringList headers_;
};

#endif  // EXPORTDSVTEST_H
