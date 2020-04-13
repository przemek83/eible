#ifndef EXPORTDSVTEST_H
#define EXPORTDSVTEST_H

#include <QObject>

class TestTableModel;

class ExportDsvTest : public QObject
{
    Q_OBJECT
private Q_SLOTS:
    void initTestCase();

    void testEmptyTableTsv();
    void testEmptyTableCsv();

    void testHeadersOnlyTsv();
    void testHeadersOnlyCsv();

    void testSimpleTableTsv();
    void testSimpleTableCsv();

    void testViewWithMultiSelectionTsv();
    void testViewWithMultiSelectionCsv();

    void testViewWithSeparatorInStringFieldTsv();
    void testViewWithSeparatorInStringFieldCsv();

    void Benchmark_data();

    void Benchmark();

    void cleanupTestCase();

private:
    void checkEmptyTable(char separator);
    void checkHeadersOnly(char separator, const QString& expected);
    void checkSimpleTable(char separator, const QString& expected);
    void checkViewWithMultiSelection(char separator, const QString& expected);
    void checkViewWithSeparatorInStringField(char separator,
                                             const QString& expected);

    static QString tableDataTsv_;
    static QString tableDataCsv_;
    static QString multiSelectionTableDataTsv_;
    static QString multiSelectionTableDataCsv_;
    static QString headersOnlyDataTsv_;
    static QString headersOnlyDataCsv_;
    static QString separatorInStringFieldDataTsv_;
    static QString separatorInStringFieldDataCsv_;
    static QString emptyData_;
    static QStringList headers_;
    TestTableModel* tableModelForBenchmarking_{nullptr};
};

#endif  // EXPORTDSVTEST_H
