#pragma once

#include <QObject>

class ImportOdsTest : public QObject
{
public:
    explicit ImportOdsTest(QObject* parent = nullptr);

    Q_OBJECT
private Q_SLOTS:
    void testRetrievingSheetNames() const;
    static void testRetrievingSheetNamesFromEmptyFile();

    static void testGetColumnList_data();
    void testGetColumnList() const;
    void testSettingEmptyColumnName() const;
    void testGetColumnListTwoSheets() const;

    static void testGetColumnTypes_data();
    void testGetColumnTypes() const;

    static void testGetColumnCount_data();
    void testGetColumnCount() const;

    static void testGetRowCount_data();
    void testGetRowCount() const;

    static void testGetRowAndColumnCountViaGetColumnTypes_data();
    void testGetRowAndColumnCountViaGetColumnTypes() const;

    static void testGetData_data();
    void testGetData() const;

    static void testGetDataLimitRows_data();
    void testGetDataLimitRows() const;

    static void testGetDataExcludeColumns_data();
    void testGetDataExcludeColumns() const;
    void testGetDataExcludeInvalidColumn() const;

    static void benchmarkGetData();

    void testEmittingProgressPercentChangedEmptyFile() const;
    void testEmittingProgressPercentChangedSmallFile() const;
    static void testEmittingProgressPercentChangedBigFile();

    void testInvalidSheetName() const;

    static void testDamagedFile();

private:
    const QString testFileName_{QStringLiteral(":/res/testOds.ods")};
    const QString templateFileName_{QStringLiteral(":/res/emptyOds.ods")};
};
