#pragma once

#include <QObject>

class TestTableModel;

class ExportDsvTest : public QObject
{
    Q_OBJECT
public:
    explicit ExportDsvTest(QObject* parent = nullptr);

private Q_SLOTS:
    static void testEmptyTable_data();
    void testEmptyTable();

    void testHeadersOnly_data();
    static void testHeadersOnly();

    void testSimpleTable_data();
    static void testSimpleTable();

    void testViewWithMultiSelection_data();
    static void testViewWithMultiSelection();

    void testSpecialCharInStringField_data();
    static void testSpecialCharInStringField();

    void testCustomDateFormat();
    void testIsoShortDate();

    void testLocaleForNumbers();

    void benchmark_data();
    void benchmark();

    void cleanupTestCase();

private:
    QString tableDataTsv_{
        "Text\tNumeric\tDate\nItem 0 0\t1.00\t2020-01-03\nItem 0 "
        "1\t2.00\t2020-01-04\nItem 0 2\t3.00\t2020-01-05"};
    QString tableDataCsv_{
        "Text,Numeric,Date\nItem 0 0,1.00,2020-01-03\nItem 0 "
        "1,2.00,2020-01-04\nItem 0 2,3.00,2020-01-05"};
    QString multiSelectionTableDataTsv_{
        "Text\tNumeric\tDate\nItem 0 0\t1.00\t2020-01-03\nItem 0 "
        "2\t3.00\t2020-01-05"};
    QString multiSelectionTableDataCsv_{
        "Text,Numeric,Date\nItem 0 0,1.00,2020-01-03\nItem 0 "
        "2,3.00,2020-01-05"};
    QString headersOnlyDataTsv_{QStringLiteral("Text\tNumeric\tDate")};
    QString headersOnlyDataCsv_{QStringLiteral("Text,Numeric,Date")};
    QString separatorInStringFieldDataTsv_{
        "Text\tNumeric\tDate\n\"Other\titem\"\t1.00\t2020-01-03\nItem 0 "
        "1\t2.00\t2020-01-04"};
    QString separatorInStringFieldDataCsv_{
        "Text,Numeric,Date\n\"Other,item\",1.00,2020-01-03\nItem 0 "
        "1,2.00,2020-01-04"};
    QString newLineInStringFieldDataTsv_{
        "Text\tNumeric\tDate\n\"Other\nitem\"\t1.00\t2020-01-03\nItem 0 "
        "1\t2.00\t2020-01-04"};
    QString newLineInStringFieldDataCsv_{
        "Text,Numeric,Date\n\"Other\nitem\",1.00,2020-01-03\nItem 0 "
        "1,2.00,2020-01-04"};
    QString doubleQuotesInStringFieldDataTsv_{
        "Text\tNumeric\tDate\n\"Other\"\"item\"\t1.00\t2020-01-03\nItem 0 "
        "1\t2.00\t2020-01-04"};
    QString doubleQuotesInStringFieldDataCsv_{
        "Text,Numeric,Date\n\"Other\"\"item\",1.00,2020-01-03\nItem 0 "
        "1,2.00,2020-01-04"};
    QString customDateFormatData_{
        QStringLiteral("Text,Numeric,Date\nItem 0 0,1.00,03/01/20")};
    QString isoShortDateData_{
        QStringLiteral("Text,Numeric,Date\nItem 0 0,1.00,2020-01-03")};
    QString localeForNumbersData_{
        QStringLiteral("Text\tNumeric\tDate\nItem 0 0\t1,00\t2020-01-03")};
    QString emptyData_{QLatin1String("")};
    QStringList headers_{QStringLiteral("Text"), QStringLiteral("Numeric"),
                         QStringLiteral("Date")};

    TestTableModel* tableModelForBenchmarking_{nullptr};
};
