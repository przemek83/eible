#ifndef EXPORTDSVTEST_H
#define EXPORTDSVTEST_H

#include <QObject>

class TestTableModel;

class ExportDsvTest : public QObject
{
    Q_OBJECT
private Q_SLOTS:
    void initTestCase();

    void testExportingEmptyTableTsv();
    void testExportingEmptyTableCsv();

    void testExportingHeadersOnlyTsv();
    void testExportingHeadersOnlyCsv();

    void testExportingSimpleTableTsv();
    void testExportingSimpleTableCsv();

    void testExportingViewWithMultiSelectionTsv();
    void testExportingViewWithMultiSelectionCsv();

    void Benchmark_data();

    void Benchmark();

    void cleanupTestCase();

private:
    void checkExportingEmptyTable(char separator);
    void checkExportingHeadersOnly(char separator, const QString& expected);
    void checkExportingSimpleTable(char separator, const QString& expected);
    void checkExportingViewWithMultiSelection(char separator,
                                              const QString& expected);

    static QString tableDataTsv_;
    static QString tableDataCsv_;
    static QString multiSelectionTableDataTsv_;
    static QString multiSelectionTableDataCsv_;
    static QString headersOnlyDataTsv_;
    static QString headersOnlyDataCsv_;
    static QString emptyData_;
    static QStringList headers_;
    TestTableModel* tableModelForBenchmarking_{nullptr};
};

#endif  // EXPORTDSVTEST_H
