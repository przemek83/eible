#pragma once

#include <QObject>

class ImportOdsTest : public QObject
{
public:
    explicit ImportOdsTest(QObject* parent = nullptr);

    Q_OBJECT
private Q_SLOTS:
    void testRetrievingSheetNames();
    void testRetrievingSheetNamesFromEmptyFile();

    void testGetColumnList_data();
    void testGetColumnList();
    void testSettingEmptyColumnName();
    void testGetColumnListTwoSheets();

    void testGetColumnTypes_data();
    void testGetColumnTypes();

    void testGetColumnCount_data();
    void testGetColumnCount();

    void testGetRowCount_data();
    void testGetRowCount();

    void testGetRowAndColumnCountViaGetColumnTypes_data();
    void testGetRowAndColumnCountViaGetColumnTypes();

    void testGetData_data();
    void testGetData();

    void testGetDataLimitRows_data();
    void testGetDataLimitRows();

    void testGetDataExcludeColumns_data();
    void testGetDataExcludeColumns();
    void testGetDataExcludeInvalidColumn();

    void benchmarkGetData();

    void testEmittingProgressPercentChangedEmptyFile();
    void testEmittingProgressPercentChangedSmallFile();
    void testEmittingProgressPercentChangedBigFile();

    void testInvalidSheetName();

private:
    const QString testFileName_{QStringLiteral(":/testOds.ods")};
    const QString templateFileName_{QStringLiteral(":/emptyOds.ods")};
};
