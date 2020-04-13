#ifndef EXPORTDSVTEST_H
#define EXPORTDSVTEST_H

#include <QObject>

class ExportDsvTest : public QObject
{
    Q_OBJECT
private Q_SLOTS:
    void initTestCase();

    void testExportingEmptyTable();

    void testExportingHeadersOnly();

    void testExportingSimpleTable();

    void testExportingViewWithMultiSelection();

    //    void Benchmark_data();

    //    void Benchmark();

    //    void cleanupTestCase();

private:
    static QString tableSheetData_;
    static QString multiSelectionTableSheetData_;
    static QString headersOnlySheetData_;
    static QString emptySheetData_;
    static QStringList headers_;
};

#endif  // EXPORTDSVTEST_H
