#include "ImportCommon.h"

#include <eible/ImportOds.h>
#include <eible/ImportSpreadsheet.h>
#include <eible/ImportXlsx.h>
#include <QBuffer>
#include <QFile>
#include <QSignalSpy>
#include <QTest>

namespace
{
QList<QStringList> getTestColumnNames()
{
    QList<QStringList> testColumnNames{
        {"Text", "Numeric", "Date"},
        {"Trait #1", "Value #1", "Transaction date", "Units", "Price",
         "Price per unit", "Another trait"},
        {},
        {"cena nier", "pow", "cena metra", "data transakcji", "text"},
        {"name", "date", "mass (kg)", "height", "---", "---", "---", "---",
         "---", "---", "---", "---"},
        {"modificator", "x", "y"},
        {"Pow", "Cena", "cena_m", "---", "---", "---"},
        {"---", "second", "third", "fourth"},
        {"user", "pass", "lic_exp", "uwagi"}};

    return testColumnNames;
}

QVector<int> getExpectedRowCounts() { return {2, 19, 0, 4, 30, 25, 20, 3, 3}; }

QVector<QVector<ColumnType>> getColumnTypes()
{
    QVector<QVector<ColumnType>> columnTypes{
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
        {ColumnType::DATE, ColumnType::STRING, ColumnType::NUMBER,
         ColumnType::STRING},
        {ColumnType::STRING, ColumnType::STRING, ColumnType::STRING,
         ColumnType::STRING}};
    return columnTypes;
}

QVector<QVector<QVector<QVariant>>> getSheetData()
{
    QVector<QVector<QVector<QVariant>>> sheetData{
        {{"Item 0, 0", 1., QDate(2020, 1, 3)},
         {"Other item", 2., QDate(2020, 1, 4)}},
        {{"trait 1", 1., QDate(2010, 1, 7), 48.49, 550060.8, 11343.8, "brown"},
         {"trait 2", 2., QDate(2010, 1, 8), 45.98, 236341.18, 5140.09, "red"},
         {"trait 3", 9., QDate(2010, 1, 9), 51.33, 422955.71, 8239.93,
          "yellow"},
         {"trait 4", 10., QDate(2010, 1, 10), 83.58, 372950.9, 4462.2, "black"},
         {"trait 16", 22., QDate(2010, 1, 22), 33.06, 281118.11, 8503.27,
          "yellow"},
         {"trait 17", 24., QDate(2010, 1, 24), 61.22, 352779.19, 5762.48,
          "black"},
         {"trait 18", 25., QDate(2010, 1, 25), 81.16, 403985.68, 4977.65,
          "white"},
         {"trait 19", 27., QDate(2010, 1, 27), 78.16, 410101.03, 5246.94,
          "black"},
         {"trait 20", 28., QDate(2010, 1, 28), 73.66, 323727.84, 4394.89,
          "yellow"},
         {"trait 21", 29., QDate(2010, 1, 29), 94.68, 548623.75, 5794.51,
          "pink"},
         {"trait 22", 30., QDate(2010, 1, 30), 71.55, 413429.91, 5778.2,
          "pink"},
         {"trait 1", 31., QDate(2010, 1, 31), 80.52, 324537.82, 4030.52,
          "blue"},
         {"trait 2", 32., QDate(2010, 2, 1), 92.71, 473075.07, 5102.74,
          "black"},
         {"trait 3", 33., QDate(2010, 2, 2), 58.86, 339724.62, 5771.74, "blue"},
         {"trait 4", 34., QDate(2010, 2, 3), 70.17, 443233.11, 6316.56, "pink"},
         {"trait 5", 35., QDate(2010, 2, 4), 82.71, 426391.14, 5155.25,
          "black"},
         {"trait 6", 36., QDate(2010, 2, 5), 60.89, 425907.85, 6994.71, "blue"},
         {"trait 7", 37., QDate(2010, 2, 6), 64.22, 444124.57, 6915.67, "blue"},
         {"trait 8", 38., QDate(2010, 2, 7), 68.32, 353874.01, 5179.65,
          "brown"}},
        {},
        {{200000., 51., 3921.56862745098, QDate(2012, 02, 02), "a"},
         {200001., 52., 3846.17307692308, QDate(), "b"},
         {200002., 53., 3773.62264150943, QDate(2012, 02, 04), "b"},
         {200003., 54., 3703.75925925926, QDate(2012, 02, 05), "c"}},
        {{QVariant(QMetaType(QMetaType::QString)), QDate(1970, 1, 1), 52.21,
          1.47, QDate(1970, 1, 1), 0., "Sx", "105", 24.76, 0., 2725.8841, 0},
         {QVariant(QMetaType(QMetaType::QString)), QDate(1970, 1, 2), 53.12,
          1.5, QVariant(QMetaType(QMetaType::QDate)), 1., "Sy", "931.17",
          931.17, 1., 2821.7344, 53.12},
         {QVariant(QMetaType(QMetaType::QString)), QDate(1970, 1, 3), 54.48,
          1.52, QVariant(QMetaType(QMetaType::QDate)), 2., "Sxx", "1015",
          41.0532, 4., 2968.0704, 108.96},
         {QVariant(QMetaType(QMetaType::QString)), QDate(1970, 1, 4), 55.84,
          1.55, QVariant(QMetaType(QMetaType::QDate)), 3., "Syy", "58498.5439",
          58498.5439, 9., 3118.1056, 167.52},
         {QVariant(QMetaType(QMetaType::QString)), QDate(1970, 1, 5), 57.2,
          1.57, QVariant(QMetaType(QMetaType::QDate)), 4., "Sxy", "6956.82",
          1548.2453, 16., 3271.84, 228.8},
         {QVariant(QMetaType(QMetaType::QString)), QDate(1970, 1, 6), 58.57,
          1.6, QVariant(QMetaType(QMetaType::QDate)), 5., "n", "15", 15, 25.,
          3430.4449, 292.85},
         {QVariant(QMetaType(QMetaType::QString)), QDate(1970, 1, 7), 59.93,
          1.63, QVariant(QMetaType(QMetaType::QDate)), 6., "a (beta w wiki)",
          "1.56653571428571", 61.2721865421088, 36., 3591.6049, 359.58},
         {QVariant(QMetaType(QMetaType::QString)), QDate(1970, 1, 8), 61.29,
          1.65, QVariant(QMetaType(QMetaType::QDate)), 7., "b (alfa w wiki)",
          "51.11225", -39.0619559188409, 49., 3756.4641, 429.03},
         {QVariant(QMetaType(QMetaType::QString)), QDate(1970, 1, 9), 63.11,
          1.68, QVariant(QMetaType(QMetaType::QDate)), 8.,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0., 64., 3982.8721, 504.88},
         {QVariant(QMetaType(QMetaType::QString)), QDate(1970, 1, 10), 64.47,
          1.7, QVariant(QMetaType(QMetaType::QDate)), 9.,
          QVariant(QMetaType(QMetaType::QString)), "y=ax+b", 0., 81., 4156.3809,
          580.23},
         {QVariant(QMetaType(QMetaType::QString)), QDate(1970, 1, 11), 66.28,
          1.73, QVariant(QMetaType(QMetaType::QDate)), 10.,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0., 100., 4393.0384, 662.8},
         {QVariant(QMetaType(QMetaType::QString)), QDate(1970, 1, 12), 68.1,
          1.75, QVariant(QMetaType(QMetaType::QDate)), 11.,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0., 121., 4637.61, 749.1},
         {QVariant(QMetaType(QMetaType::QString)), QDate(1970, 1, 13), 69.92,
          1.78, QVariant(QMetaType(QMetaType::QDate)), 12.,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0., 144., 4888.8064, 839.04},
         {QVariant(QMetaType(QMetaType::QString)), QDate(1970, 1, 14), 72.19,
          1.8, QVariant(QMetaType(QMetaType::QDate)), 13.,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0., 169., 5211.3961, 938.47},
         {QVariant(QMetaType(QMetaType::QString)), QDate(1970, 1, 15), 74.46,
          1.83, QVariant(QMetaType(QMetaType::QDate)), 14.,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0., 196., 5544.2916,
          1042.44},
         {QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QDate)), 0., 0.,
          QVariant(QMetaType(QMetaType::QDate)), 0.,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0., 2.1609, 2725.8841,
          76.7487},
         {QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QDate)), 0., 0.,
          QVariant(QMetaType(QMetaType::QDate)), 0.,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0., 2.25, 2821.7344, 79.68},
         {QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QDate)), 0., 0.,
          QVariant(QMetaType(QMetaType::QDate)), 0.,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0., 2.3104, 2968.0704,
          82.8096},
         {QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QDate)), 0., 0.,
          QVariant(QMetaType(QMetaType::QDate)), 0.,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0., 2.4025, 3118.1056,
          86.552},
         {QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QDate)), 0., 0.,
          QVariant(QMetaType(QMetaType::QDate)), 0.,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0., 2.4649, 3271.84, 89.804},
         {QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QDate)), 0., 0.,
          QVariant(QMetaType(QMetaType::QDate)), 0.,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0., 2.56, 3430.4449, 93.712},
         {QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QDate)), 0., 0.,
          QVariant(QMetaType(QMetaType::QDate)), 0.,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0., 2.6569, 3591.6049,
          97.6859},
         {QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QDate)), 0., 0.,
          QVariant(QMetaType(QMetaType::QDate)), 0.,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0., 2.7225, 3756.4641,
          101.1285},
         {QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QDate)), 0., 0.,
          QVariant(QMetaType(QMetaType::QDate)), 0.,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0., 2.8224, 3982.8721,
          106.0248},
         {QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QDate)), 0., 0.,
          QVariant(QMetaType(QMetaType::QDate)), 0.,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0., 2.89, 4156.3809,
          109.599},
         {QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QDate)), 0., 0.,
          QVariant(QMetaType(QMetaType::QDate)), 0.,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0., 2.9929, 4393.0384,
          114.6644},
         {QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QDate)), 0., 0.,
          QVariant(QMetaType(QMetaType::QDate)), 0.,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0., 3.0625, 4637.61,
          119.175},
         {QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QDate)), 0., 0.,
          QVariant(QMetaType(QMetaType::QDate)), 0.,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0., 3.1684, 4888.8064,
          124.4576},
         {QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QDate)), 0., 0.,
          QVariant(QMetaType(QMetaType::QDate)), 0.,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0., 3.24, 5211.3961,
          129.942},
         {QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QDate)), 0., 0.,
          QVariant(QMetaType(QMetaType::QDate)), 0.,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0., 3.3489, 5544.2916,
          136.2618}},
        {{0., 5.6494487, 0.0795697},  {0.1, 5.9721082, 0.0841142},
         {0.2, 6.3006678, 0.0887418}, {0.3, 6.634098, 0.093438},
         {0.4, 6.9713125, 0.0981875}, {0.5, 7.311154, 0.102974},
         {0.6, 7.652309, 0.107779},   {0.7, 7.993464, 0.112584},
         {0.8, 8.33327, 0.11737},     {0.9, 8.670307, 0.122117},
         {1., 9.003013, 0.126803},    {1.1, 9.329968, 0.131408},
         {1.2, 9.649539, 0.135909},   {1.3, 9.960306, 0.140286},
         {1.4, 10.260636, 0.144516},  {1.5, 10.549109, 0.148579},
         {1.6, 10.824092, 0.152452},  {1.7, 11.084236, 0.156116},
         {1.8, 11.32805, 0.15955},    {1.9, 11.554327, 0.162737},
         {2., 11.761647, 0.165657},   {2.1, 11.948945, 0.168295},
         {2.2, 12.115156, 0.170636},  {2.3, 12.259215, 0.172665},
         {2.4, 12.380412, 0.174372}},
        {{56., 423000., 7553.57142857143,
          QVariant(QMetaType(QMetaType::QString)), "Średnia:", 6715.2873295802},
         {45., 345000., 7666.66666666667,
          QVariant(QMetaType(QMetaType::QString)),
          "Mediana:", 6277.77777777778},
         {76., 455664., 5995.57894736842,
          QVariant(QMetaType(QMetaType::QString)), "Q10", 4301.1655011655},
         {66., 330000., 5000., QVariant(QMetaType(QMetaType::QString)), "Q25",
          5065.78947368421},
         {46., 450000., 9782.60869565217,
          QVariant(QMetaType(QMetaType::QString)), "Q50", 6277.77777777778},
         {65., 280000., 4307.69230769231,
          QVariant(QMetaType(QMetaType::QString)), "Q75", 7668.62777777778},
         {65., 250000., 3846.15384615385,
          QVariant(QMetaType(QMetaType::QString)), "Q90", 9856.97940503433},
         {60., 380000., 6333.33333333333,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0.},
         {55., 450000., 8181.81818181818,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0.},
         {76., 800000., 10526.3157894737,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0.},
         {57., 290000., 5087.71929824562,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0.},
         {67., 410000., 6119.40298507463,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0.},
         {41., 310000., 7560.9756097561,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0.},
         {68., 390000., 5735.29411764706,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0.},
         {59., 400000., 6779.66101694915,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0.},
         {66., 280000., 4242.42424242424,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0.},
         {45., 280000., 6222.22222222222,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0.},
         {55., 270000., 4909.09090909091,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0.},
         {45., 345353., 7674.51111111111,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0.},
         {34., 366544., 10780.7058823529,
          QVariant(QMetaType(QMetaType::QString)),
          QVariant(QMetaType(QMetaType::QString)), 0.}},
        {{QDate(2012, 1, 1), "a", 45., QVariant(QMetaType(QMetaType::QString))},
         {QVariant(QMetaType(QMetaType::QDate)), "s", 5.,
          QVariant(QMetaType(QMetaType::QString))},
         {QVariant(QMetaType(QMetaType::QDate)), "d", 67.,
          QVariant(QMetaType(QMetaType::QString))}},
        {{"test1", "test123", "zawsze aktualna",
          QVariant(QMetaType(QMetaType::QString))},
         {"test2", "test123", QVariant(QMetaType(QMetaType::QString)),
          "konto zablokowane"},
         {"test3", "test123", "2011-09-30",
          QVariant(QMetaType(QMetaType::QString))}}};
    return sheetData;
}

constexpr int NO_SIGNAL{0};
}  // namespace

namespace ImportCommon
{
void checkRetrievingSheetNames(ImportSpreadsheet& importer)
{
    auto [success, actualSheetNames]{importer.getSheetNames()};
    QCOMPARE(success, true);
    QCOMPARE(actualSheetNames, getSheetNames());
}

void checkRetrievingSheetNamesFromEmptyFile(ImportSpreadsheet& importer)
{
    auto [success, actualNames]{importer.getSheetNames()};
    QCOMPARE(success, false);
    QCOMPARE(actualNames, {});
}

void prepareDataForGetColumnListTest()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<QStringList>("expectedColumnList");
    const QList<QStringList> testColumnNames{getTestColumnNames()};
    for (int i{0}; i < testColumnNames.size(); ++i)
    {
        const QString& sheetName{getSheetNames()[i]};
        QTest::newRow(("Columns in " + sheetName).toStdString().c_str())
            << sheetName << testColumnNames[i];
    }
}

void checkGetColumnList(ImportSpreadsheet& importer)
{
    QFETCH(const QString, sheetName);
    QFETCH(const QStringList, expectedColumnList);
    auto [success, actualColumnList]{importer.getColumnNames(sheetName)};
    QCOMPARE(success, true);
    QCOMPARE(actualColumnList, expectedColumnList);
}

void prepareDataForGetColumnTypes()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<QVector<ColumnType>>("expectedColumnTypes");

    const qsizetype testColumnNamesSize{getTestColumnNames().size()};
    const QVector<QVector<ColumnType>> columnTypes{getColumnTypes()};
    for (qsizetype i{0}; i < testColumnNamesSize; ++i)
    {
        const QString& sheetName{getSheetNames()[i]};
        QTest::newRow(("Columns types in " + sheetName).toStdString().c_str())
            << sheetName << columnTypes[i];
    }
}

void checkGetColumnTypes(ImportSpreadsheet& importer)
{
    QFETCH(const QString, sheetName);
    QFETCH(const QVector<ColumnType>, expectedColumnTypes);
    auto [success, columnTypes]{importer.getColumnTypes(sheetName)};
    QCOMPARE(success, true);
    QCOMPARE(columnTypes, expectedColumnTypes);
}

void checkSettingEmptyColumnName(ImportSpreadsheet& importer)
{
    const QString newEmptyColumnName{QStringLiteral("<empty column>")};
    importer.setNameForEmptyColumn(newEmptyColumnName);
    auto [success,
          actualColumnList]{importer.getColumnNames(getSheetNames()[4])};

    const QStringList columnNames{getTestColumnNames().at(4)};
    std::list<QString> expectedColumnList(
        static_cast<std::size_t>(columnNames.size()));
    std::replace_copy(columnNames.constBegin(), columnNames.constEnd(),
                      expectedColumnList.begin(), QStringLiteral("---"),
                      newEmptyColumnName);

    QCOMPARE(success, true);
    QCOMPARE(actualColumnList,
             QList(expectedColumnList.begin(), expectedColumnList.end()));
}

void checkGetColumnListTwoSheets(ImportSpreadsheet& importer)
{
    auto [success,
          actualColumnList]{importer.getColumnNames(getSheetNames()[4])};
    QCOMPARE(success, true);
    QList<QStringList> columnNames{getTestColumnNames()};
    QCOMPARE(actualColumnList, columnNames[4]);

    std::tie(success, actualColumnList) =
        importer.getColumnNames(getSheetNames()[0]);
    QCOMPARE(success, true);
    QCOMPARE(actualColumnList, columnNames[0]);
}

void prepareDataForGetColumnCountTest()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<int>("expectedColumnCount");

    const QList<QStringList> testColumnNames{getTestColumnNames()};
    for (int i{0}; i < testColumnNames.size(); ++i)
    {
        const QString& sheetName{getSheetNames()[i]};
        QTest::newRow(("Columns in " + sheetName).toStdString().c_str())
            << sheetName << testColumnNames[i].size();
    }
}

void checkGetColumnCount(ImportSpreadsheet& importer)
{
    QFETCH(const QString, sheetName);
    QFETCH(const int, expectedColumnCount);
    auto [success, actualColumnCount]{importer.getColumnCount(sheetName)};
    QCOMPARE(success, true);
    QCOMPARE(actualColumnCount, expectedColumnCount);
}

void prepareDataForGetRowCountTest()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<int>("expectedRowCount");

    const qsizetype testColumnNamesSize{getTestColumnNames().size()};
    QVector<int> expectedRowCounts{getExpectedRowCounts()};
    for (qsizetype i{0}; i < testColumnNamesSize; ++i)
    {
        const QString& sheetName{getSheetNames()[i]};
        QTest::newRow(("Rows in " + sheetName).toStdString().c_str())
            << sheetName << expectedRowCounts[i];
    }
}

void checkGetRowCount(ImportSpreadsheet& importer)
{
    QFETCH(const QString, sheetName);
    QFETCH(const int, expectedRowCount);
    auto [success, actualRowCount]{importer.getRowCount(sheetName)};
    QCOMPARE(success, true);
    QCOMPARE(actualRowCount, expectedRowCount);
}

void prepareDataForGetRowAndColumnCountViaGetColumnTypes()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<int>("expectedRowCount");
    QTest::addColumn<int>("expectedColumnCount");

    const QList<QStringList> testColumnNames{getTestColumnNames()};
    QVector<int> expectedRowCounts{getExpectedRowCounts()};
    for (int i{0}; i < testColumnNames.size(); ++i)
    {
        const QString& sheetName{getSheetNames()[i]};
        QTest::newRow(("Rows and columns via GetColumnTypes() in " + sheetName)
                          .toStdString()
                          .c_str())
            << sheetName << expectedRowCounts[i]
            << static_cast<int>(testColumnNames[i].size());
    }
}

void testGetRowAndColumnCountViaGetColumnTypes(ImportSpreadsheet& importer)
{
    QFETCH(const QString, sheetName);
    QFETCH(const int, expectedRowCount);
    QFETCH(const int, expectedColumnCount);
    importer.getColumnTypes(sheetName);
    auto [successRowCount, actualRowCount]{importer.getRowCount(sheetName)};
    QCOMPARE(successRowCount, true);
    QCOMPARE(actualRowCount, expectedRowCount);
    auto [successColumnCount,
          actualColumnCount]{importer.getColumnCount(sheetName)};
    QCOMPARE(successColumnCount, true);
    QCOMPARE(actualColumnCount, expectedColumnCount);
}

void prepareDataForGetData()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<QVector<QVector<QVariant>>>("expectedData");

    const QVector<QVector<QVector<QVariant>>> sheetData{getSheetData()};
    for (int i{0}; i < sheetData.size(); ++i)
    {
        const QString& sheetName{getSheetNames()[i]};
        QTest::newRow(("Data in " + sheetName).toStdString().c_str())
            << sheetName << sheetData[i];
    }
}

void checkGetData(ImportSpreadsheet& importer)
{
    QFETCH(const QString, sheetName);
    QFETCH(const QVector<QVector<QVariant>>, expectedData);
    auto [success, actualData]{importer.getData(sheetName, {})};
    QCOMPARE(success, true);
    QCOMPARE(actualData, expectedData);
}

void prepareDataForGetDataLimitRows()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<int>("rowLimit");
    QTest::addColumn<QVector<QVector<QVariant>>>("expectedData");
    QString sheetName{getSheetNames()[0]};

    const QVector<QVector<QVector<QVariant>>> sheetData{getSheetData()};
    QTest::newRow(("Limited data to 10 in " + sheetName).toStdString().c_str())
        << sheetName << 10 << sheetData[0];
    QTest::newRow(("Limited data to 2 in " + sheetName).toStdString().c_str())
        << sheetName << 2 << sheetData[0];
    sheetName = getSheetNames()[5];
    QVector<QVector<QVariant>> expectedValues;
    const int rowLimit{12};
    expectedValues.reserve(rowLimit);
    for (int i{0}; i < rowLimit; ++i)
        expectedValues.append(sheetData[5][i]);
    QTest::newRow(
        ("Limited data to " + QString::number(rowLimit) + " in " + sheetName)
            .toStdString()
            .c_str())
        << sheetName << rowLimit << expectedValues;
}

void checkGetDataLimitRows(ImportSpreadsheet& importer)
{
    QFETCH(const QString, sheetName);
    QFETCH(const int, rowLimit);
    QFETCH(const QVector<QVector<QVariant>>, expectedData);
    auto [success,
          actualData]{importer.getLimitedData(sheetName, {}, rowLimit)};
    QCOMPARE(success, true);
    QCOMPARE(actualData, expectedData);
}

QVector<QVector<QVariant>> getDataWithoutColumns(
    const QVector<QVector<QVariant>>& data,
    const QVector<int>& columnsToExclude)
{
    QVector<QVector<QVariant>> expectedValues;
    expectedValues.reserve(data.size());
    for (auto dataRow : data)
    {
        for (const int column : columnsToExclude)
            dataRow.remove(column);
        expectedValues.append(dataRow);
    }
    return expectedValues;
}

void prepareDataForGetGetDataExcludeColumns()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<QVector<int>>("excludedColumns");
    QTest::addColumn<QVector<QVector<QVariant>>>("expectedData");

    const QString sheetName{getSheetNames()[0]};
    QString testName{"Get data with excluded column 1 in " + sheetName};

    const QVector<QVector<QVector<QVariant>>> sheetData{getSheetData()};
    QTest::newRow(qUtf8Printable(testName))
        << sheetName << QVector<int>{1}
        << getDataWithoutColumns(sheetData[0], {1});

    testName = "Get data with excluded column 0 and 2 in " + sheetName;
    QTest::newRow(qUtf8Printable(testName))
        << sheetName << QVector<int>{0, 2}
        << getDataWithoutColumns(sheetData[0], {2, 0});
}

void checkGetDataExcludeColumns(ImportSpreadsheet& importer)
{
    QFETCH(const QString, sheetName);
    QFETCH(const QVector<int>, excludedColumns);
    QFETCH(const QVector<QVector<QVariant>>, expectedData);
    auto [success, actualData]{importer.getData(sheetName, excludedColumns)};
    QCOMPARE(success, true);
    QCOMPARE(actualData, expectedData);
}

void checkGetDataExcludeInvalidColumn(ImportSpreadsheet& importer)
{
    auto [success, actualData]{importer.getData(getSheetNames()[0], {3})};
    QCOMPARE(success, false);
}

QVector<QVector<QVariant>> getDataForSheet(const QString& fileName)
{
    const QVector<QVector<QVector<QVariant>>> sheetData{getSheetData()};
    return sheetData[getSheetNames().indexOf(fileName)];
}

QStringList& getSheetNames()
{
    static QStringList sheetNames_{
        QStringLiteral("Sheet1"),        QStringLiteral("Sheet2"),
        QStringLiteral("Sheet3(empty)"), QStringLiteral("Sheet4"),
        QStringLiteral("Sheet5"),        QStringLiteral("Sheet6"),
        QStringLiteral("Sheet7"),        QStringLiteral("Sheet9"),
        QStringLiteral("testAccounts")};
    return sheetNames_;
}

void checkInvalidSheetName(ImportSpreadsheet& importer)
{
    auto [success, actualSharedStrings]{
        importer.getColumnCount(QStringLiteral("invalidSheetName"))};
    QCOMPARE(success, false);
}

void checkEmittingProgressPercentChangedEmptyFile(ImportSpreadsheet& importer)
{
    const QSignalSpy spy(&importer, &ImportSpreadsheet::progressPercentChanged);
    QStringList sheetNames{importer.getSheetNames().second};
    auto [success, actualData]{importer.getData(sheetNames.first(), {})};
    QCOMPARE(success, true);
    QCOMPARE(spy.count(), NO_SIGNAL);
}

void checkEmittingProgressPercentChangedSmallFile(ImportSpreadsheet& importer)
{
    const QSignalSpy spy(&importer, &ImportSpreadsheet::progressPercentChanged);
    const int sheetIndex{1};
    auto [success,
          actualData]{importer.getData(getSheetNames()[sheetIndex], {})};
    QCOMPARE(success, true);
    QCOMPARE(spy.count(), 19);
}

void checkEmittingProgressPercentChangedBigFile(ImportSpreadsheet& importer)
{
    const QSignalSpy spy(&importer, &ImportSpreadsheet::progressPercentChanged);
    QStringList sheetNames{importer.getSheetNames().second};
    auto [success, actualData]{importer.getData(sheetNames.first(), {})};
    QCOMPARE(success, true);
    QCOMPARE(spy.count(), 100);
}

}  // namespace ImportCommon
