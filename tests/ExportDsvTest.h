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

    void testExportingHeadersOnlyTsv();
    void testExportingHeadersOnlyCsv();

    void testExportingSimpleTableTsv();
    void testExportingSimpleTableCsv();

    void testExportingViewWithMultiSelection();

    //    void Benchmark_data();

    //    void Benchmark();

    //    void cleanupTestCase();

private:
    void checkExportingEmptyTable(char separator);
    void checkExportingHeadersOnly(char separator, const QString& expected);
    void checkExportingSimpleTable(char separator, const QString& expected);

    static QString tableDataTsv_;
    static QString tableDataCsv_;
    static QString multiSelectionTableData_;
    static QString headersOnlyDataTsv_;
    static QString headersOnlyDataCsv_;
    static QString emptyData_;
    static QStringList headers_;
};

#endif  // EXPORTDSVTEST_H
