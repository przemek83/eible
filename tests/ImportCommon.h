#ifndef IMPORTCOMMON_H
#define IMPORTCOMMON_H

#include <ImportSpreadsheet.h>
#include <QStringList>

class ImportCommon
{
public:
    template <class T>
    static void checkRetrievingSheetNames(const QString& fileName);

    template <class T>
    static void checkRetrievingSheetNamesFromEmptyFile();

    static void prepareDataForGetColumnListTest();
    template <class T>
    static void checkGetColumnList(const QString& fileName);

    template <class T>
    static void checkSettingEmptyColumnName(const QString& fileName);

    template <class T>
    static void checkGetColumnListTwoSheets(const QString& fileName);

    static void prepareDataForGetColumnTypes();
    template <class T>
    static void checkGetColumnTypes(const QString& fileName);

    static void prepareDataForGetColumnCountTest();
    template <class T>
    static void checkGetColumnCount(const QString& fileName);

    static void prepareDataForGetRowCountTest();
    template <class T>
    static void checkGetRowCount(const QString& fileName);

private:
    static const QStringList sheetNames_;
    static const QList<QStringList> testColumnNames_;
    static const std::vector<unsigned int> expectedRowCounts_;
    static const QVector<QVector<ColumnType>> columnTypes_;
};

#endif  // IMPORTCOMMON_H
