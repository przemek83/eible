#ifndef EXPORTDSVTEST_H
#define EXPORTDSVTEST_H

#include <QObject>

class TestTableModel;

class ExportDsvTest : public QObject
{
    Q_OBJECT
public:
    explicit ExportDsvTest(QObject* parent = nullptr);

private Q_SLOTS:
    void testEmptyTable_data();
    void testEmptyTable();

    void testHeadersOnly_data();
    void testHeadersOnly();

    void testSimpleTable_data();
    void testSimpleTable();

    void testViewWithMultiSelection_data();
    void testViewWithMultiSelection();

    void testSpecialCharInStringField_data();
    void testSpecialCharInStringField();

    void testCustomDateFormat();
    void testDefaultLocaleShortDate();

    void testLocaleForNumbers();

    void benchmark_data();
    void benchmark();

    void cleanupTestCase();

private:
    static QString tableDataTsv_;
    static QString tableDataCsv_;
    static QString multiSelectionTableDataTsv_;
    static QString multiSelectionTableDataCsv_;
    static QString headersOnlyDataTsv_;
    static QString headersOnlyDataCsv_;
    static QString separatorInStringFieldDataTsv_;
    static QString separatorInStringFieldDataCsv_;
    static QString newLineInStringFieldDataTsv_;
    static QString newLineInStringFieldDataCsv_;
    static QString doubleQuotesInStringFieldDataTsv_;
    static QString doubleQuotesInStringFieldDataCsv_;
    static QString customDateFormatData_;
    static QString defaultLocaleShortDateData_;
    static QString localeForNumbersData_;
    static QString emptyData_;
    static QStringList headers_;
    TestTableModel* tableModelForBenchmarking_{nullptr};
};

#endif  // EXPORTDSVTEST_H
