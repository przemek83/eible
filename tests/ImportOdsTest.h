#pragma once

#include <QObject>

class ImportOdsTest : public QObject
{
public:
    explicit ImportOdsTest(QObject* parent = nullptr);

    Q_OBJECT
private Q_SLOTS:
    void testRetrievingSheetNames();
    static void testRetrievingSheetNamesFromEmptyFile();

    static void testGetColumnList_data();
    void testGetColumnList();
    void testSettingEmptyColumnName();
    void testGetColumnListTwoSheets();

    static void testGetColumnTypes_data();
    void testGetColumnTypes();

    static void testGetColumnCount_data();
    void testGetColumnCount();

    static void testGetRowCount_data();
    void testGetRowCount();

    static void testGetRowAndColumnCountViaGetColumnTypes_data();
    void testGetRowAndColumnCountViaGetColumnTypes();

    static void testGetData_data();
    void testGetData();

    static void testGetDataLimitRows_data();
    void testGetDataLimitRows();

    static void testGetDataExcludeColumns_data();
    void testGetDataExcludeColumns();
    void testGetDataExcludeInvalidColumn();

    static void benchmarkGetData();

    void testEmittingProgressPercentChangedEmptyFile();
    void testEmittingProgressPercentChangedSmallFile();
    static void testEmittingProgressPercentChangedBigFile();

    void testInvalidSheetName();

    static void testDamagedFile();

private:
    const QString testFileName_{QStringLiteral(":/testOds.ods")};
    const QString templateFileName_{QStringLiteral(":/emptyOds.ods")};
};
