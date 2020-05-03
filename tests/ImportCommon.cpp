#include "ImportCommon.h"

#include <ImportOds.h>
#include <ImportXlsx.h>
#include <QBuffer>
#include <QFile>
#include <QTest>

#include "ImportSpreadsheet.h"

const QStringList ImportCommon::sheetNames_{
    "Sheet1", "Sheet2", "Sheet3(empty)", "Sheet4",
    "Sheet5", "Sheet6", "Sheet7",        "testAccounts"};

const QList<QStringList> ImportCommon::testColumnNames_ = {
    {"Text", "Numeric", "Date"},
    {"Trait #1", "Value #1", "Transaction date", "Units", "Price",
     "Price per unit", "Another trait"},
    {},
    {"cena nier", "pow", "cena metra", "data transakcji", "text"},
    {"name", "date", "mass (kg)", "height", "---", "---", "---", "---", "---",
     "---", "---", "---"},
    {"modificator", "x", "y"},
    {"Pow", "Cena", "cena_m", "---", "---", "---"},
    {"user", "pass", "lic_exp", "uwagi"}};

const std::vector<unsigned int> ImportCommon::expectedRowCounts_{2,  19, 0,  4,
                                                                 30, 25, 20, 3};

const QVector<QVector<ColumnType>> ImportCommon::columnTypes_ = {
    {ColumnType::STRING, ColumnType::NUMBER, ColumnType::DATE},
    {ColumnType::STRING, ColumnType::NUMBER, ColumnType::DATE,
     ColumnType::NUMBER, ColumnType::NUMBER, ColumnType::NUMBER,
     ColumnType::STRING},
    {},
    {ColumnType::NUMBER, ColumnType::NUMBER, ColumnType::NUMBER,
     ColumnType::DATE, ColumnType::STRING},
    {ColumnType::STRING, ColumnType::DATE, ColumnType::NUMBER,
     ColumnType::NUMBER, ColumnType::DATE, ColumnType::NUMBER,
     ColumnType::STRING, ColumnType::STRING, ColumnType::NUMBER,
     ColumnType::NUMBER, ColumnType::NUMBER, ColumnType::NUMBER},
    {ColumnType::NUMBER, ColumnType::NUMBER, ColumnType::NUMBER},
    {ColumnType::NUMBER, ColumnType::NUMBER, ColumnType::NUMBER,
     ColumnType::STRING, ColumnType::STRING, ColumnType::NUMBER},
    {ColumnType::STRING, ColumnType::STRING, ColumnType::STRING,
     ColumnType::STRING}};

template <class T>
void ImportCommon::checkRetrievingSheetNames(const QString& fileName)
{
    QFile testFile(fileName);
    T importer(testFile);
    auto [success, actualSheetNames] = importer.getSheetNames();
    QCOMPARE(success, true);
    QCOMPARE(actualSheetNames, sheetNames_);
}
template void ImportCommon::checkRetrievingSheetNames<ImportXlsx>(
    const QString& fileName);
template void ImportCommon::checkRetrievingSheetNames<ImportOds>(
    const QString& fileName);

template <class T>
void ImportCommon::checkRetrievingSheetNamesFromEmptyFile()
{
    QByteArray byteArray;
    QBuffer emptyBuffer(&byteArray);
    T importer(emptyBuffer);
    auto [success, actualNames] = importer.getSheetNames();
    QCOMPARE(success, false);
    QCOMPARE(actualNames, {});
}
template void
ImportCommon::checkRetrievingSheetNamesFromEmptyFile<ImportXlsx>();
template void ImportCommon::checkRetrievingSheetNamesFromEmptyFile<ImportOds>();

void ImportCommon::prepareDataForGetColumnListTest()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<QStringList>("expectedColumnList");
    for (int i = 0; i < testColumnNames_.size(); ++i)
    {
        const QString& sheetName{sheetNames_[i]};
        QTest::newRow(("Columns in " + sheetName).toStdString().c_str())
            << sheetName << testColumnNames_[i];
    }
}

void ImportCommon::prepareDataForGetColumnTypes()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<QVector<ColumnType>>("expectedColumnTypes");

    for (int i = 0; i < testColumnNames_.size(); ++i)
    {
        const QString& sheetName{sheetNames_[i]};
        QTest::newRow(("Columns types in " + sheetName).toStdString().c_str())
            << sheetName << columnTypes_[i];
    }
}

template <class T>
void ImportCommon::checkGetColumnTypes(const QString& fileName)
{
    QFETCH(QString, sheetName);
    QFETCH(QVector<ColumnType>, expectedColumnTypes);

    QFile testFile(fileName);
    T importer(testFile);
    auto [success, columnTypes] = importer.getColumnTypes(sheetName);
    QCOMPARE(success, true);
    QCOMPARE(columnTypes, expectedColumnTypes);
}
template void ImportCommon::checkGetColumnTypes<ImportXlsx>(
    const QString& fileName);
template void ImportCommon::checkGetColumnTypes<ImportOds>(
    const QString& fileName);

template <class T>
void ImportCommon::checkGetColumnList(const QString& fileName)
{
    QFETCH(QString, sheetName);
    QFETCH(QStringList, expectedColumnList);
    QFile testFile(fileName);
    T importer(testFile);
    auto [success, actualColumnList] = importer.getColumnNames(sheetName);
    QCOMPARE(success, true);
    QCOMPARE(actualColumnList, expectedColumnList);
}
template void ImportCommon::checkGetColumnList<ImportXlsx>(
    const QString& fileName);
template void ImportCommon::checkGetColumnList<ImportOds>(
    const QString& fileName);

template <class T>
void ImportCommon::checkSettingEmptyColumnName(const QString& fileName)
{
    QFile testFile(fileName);
    T importer(testFile);
    const QString newEmptyColumnName{"<empty column>"};
    importer.setNameForEmptyColumn(newEmptyColumnName);
    auto [success, actualColumnList] = importer.getColumnNames(sheetNames_[4]);

    std::list<QString> expectedColumnList(testColumnNames_[4].size());
    std::replace_copy(testColumnNames_[4].begin(), testColumnNames_[4].end(),
                      expectedColumnList.begin(), QString("---"),
                      newEmptyColumnName);

    QCOMPARE(success, true);
    QCOMPARE(actualColumnList, QStringList::fromStdList(expectedColumnList));
}
template void ImportCommon::checkSettingEmptyColumnName<ImportXlsx>(
    const QString& fileName);
template void ImportCommon::checkSettingEmptyColumnName<ImportOds>(
    const QString& fileName);

template <class T>
void ImportCommon::checkGetColumnListTwoSheets(const QString& fileName)
{
    QFile testFile(fileName);
    T importer(testFile);
    auto [success, actualColumnList] = importer.getColumnNames(sheetNames_[4]);
    QCOMPARE(success, true);
    QCOMPARE(actualColumnList, testColumnNames_[4]);

    std::tie(success, actualColumnList) =
        importer.getColumnNames(sheetNames_[0]);
    QCOMPARE(success, true);
    QCOMPARE(actualColumnList, testColumnNames_[0]);
}
template void ImportCommon::checkGetColumnListTwoSheets<ImportXlsx>(
    const QString& fileName);
template void ImportCommon::checkGetColumnListTwoSheets<ImportOds>(
    const QString& fileName);

void ImportCommon::prepareDataForGetColumnCountTest()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<int>("expectedColumnCount");
    for (int i = 0; i < testColumnNames_.size(); ++i)
    {
        const QString& sheetName{sheetNames_[i]};
        QTest::newRow(("Columns in " + sheetName).toStdString().c_str())
            << sheetName << testColumnNames_[i].size();
    }
}

template <class T>
void ImportCommon::checkGetColumnCount(const QString& fileName)
{
    QFETCH(QString, sheetName);
    QFETCH(int, expectedColumnCount);

    QFile testFile(fileName);
    T importer(testFile);
    auto [success, actualColumnCount] = importer.getColumnCount(sheetName);
    QCOMPARE(success, true);
    QCOMPARE(actualColumnCount, expectedColumnCount);
}
template void ImportCommon::checkGetColumnCount<ImportXlsx>(
    const QString& fileName);
template void ImportCommon::checkGetColumnCount<ImportOds>(
    const QString& fileName);

void ImportCommon::prepareDataForGetRowCountTest()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<unsigned int>("expectedRowCount");
    for (int i = 0; i < testColumnNames_.size(); ++i)
    {
        const QString& sheetName{sheetNames_[i]};
        QTest::newRow(("Rows in " + sheetName).toStdString().c_str())
            << sheetName << expectedRowCounts_[i];
    }
}

template <class T>
void ImportCommon::checkGetRowCount(const QString& fileName)
{
    QFETCH(QString, sheetName);
    QFETCH(unsigned int, expectedRowCount);

    QFile testFile(fileName);
    T importer(testFile);
    auto [success, actualRowCount] = importer.getRowCount(sheetName);
    QCOMPARE(success, true);
    QCOMPARE(actualRowCount, expectedRowCount);
}
template void ImportCommon::checkGetRowCount<ImportXlsx>(
    const QString& fileName);
template void ImportCommon::checkGetRowCount<ImportOds>(
    const QString& fileName);

void ImportCommon::prepareDataForGetRowAndColumnCountViaGetColumnTypes()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<unsigned int>("expectedRowCount");
    QTest::addColumn<unsigned int>("expectedColumnCount");

    for (int i = 0; i < testColumnNames_.size(); ++i)
    {
        const QString& sheetName{sheetNames_[i]};
        QTest::newRow(("Rows and columns via GetColumnTypes() in " + sheetName)
                          .toStdString()
                          .c_str())
            << sheetName << expectedRowCounts_[i]
            << static_cast<unsigned int>(testColumnNames_[i].size());
    }
}

template <class T>
void ImportCommon::testGetRowAndColumnCountViaGetColumnTypes(
    const QString& fileName)
{
    QFETCH(QString, sheetName);
    QFETCH(unsigned int, expectedRowCount);
    QFETCH(unsigned int, expectedColumnCount);

    QFile testFile(fileName);
    T importer(testFile);
    importer.getColumnTypes(sheetName);
    auto [successRowCount, actualRowCount] = importer.getRowCount(sheetName);
    QCOMPARE(successRowCount, true);
    QCOMPARE(actualRowCount, expectedRowCount);
    auto [successColumnCount, actualColumnCount] =
        importer.getColumnCount(sheetName);
    QCOMPARE(successColumnCount, true);
    QCOMPARE(actualColumnCount, expectedColumnCount);
}
template void ImportCommon::testGetRowAndColumnCountViaGetColumnTypes<
    ImportXlsx>(const QString& fileName);
template void ImportCommon::testGetRowAndColumnCountViaGetColumnTypes<
    ImportOds>(const QString& fileName);
