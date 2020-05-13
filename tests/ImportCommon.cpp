#include "ImportCommon.h"

#include <ImportOds.h>
#include <ImportSpreadsheet.h>
#include <ImportXlsx.h>
#include <QBuffer>
#include <QFile>
#include <QSignalSpy>
#include <QTest>

const QStringList ImportCommon::sheetNames_{
    QStringLiteral("Sheet1"),        QStringLiteral("Sheet2"),
    QStringLiteral("Sheet3(empty)"), QStringLiteral("Sheet4"),
    QStringLiteral("Sheet5"),        QStringLiteral("Sheet6"),
    QStringLiteral("Sheet7"),        QStringLiteral("Sheet9"),
    QStringLiteral("testAccounts")};

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
    {"---", "second", "third", "fourth"},
    {"user", "pass", "lic_exp", "uwagi"}};

const QVector<unsigned int> ImportCommon::expectedRowCounts_{2,  19, 0, 4, 30,
                                                             25, 20, 3, 3};

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
    {ColumnType::DATE, ColumnType::STRING, ColumnType::NUMBER,
     ColumnType::STRING},
    {ColumnType::STRING, ColumnType::STRING, ColumnType::STRING,
     ColumnType::STRING}};

const QVector<QVector<QVector<QVariant>>> ImportCommon::sheetData_ = {
    {{"Item 0, 0", 1., QDate(2020, 1, 3)},
     {"Other item", 2., QDate(2020, 1, 4)}},
    {{"trait 1", 1., QDate(2010, 1, 7), 48.49, 550060.8, 11343.8, "brown"},
     {"trait 2", 2., QDate(2010, 1, 8), 45.98, 236341.18, 5140.09, "red"},
     {"trait 3", 9., QDate(2010, 1, 9), 51.33, 422955.71, 8239.93, "yellow"},
     {"trait 4", 10., QDate(2010, 1, 10), 83.58, 372950.9, 4462.2, "black"},
     {"trait 16", 22., QDate(2010, 1, 22), 33.06, 281118.11, 8503.27, "yellow"},
     {"trait 17", 24., QDate(2010, 1, 24), 61.22, 352779.19, 5762.48, "black"},
     {"trait 18", 25., QDate(2010, 1, 25), 81.16, 403985.68, 4977.65, "white"},
     {"trait 19", 27., QDate(2010, 1, 27), 78.16, 410101.03, 5246.94, "black"},
     {"trait 20", 28., QDate(2010, 1, 28), 73.66, 323727.84, 4394.89, "yellow"},
     {"trait 21", 29., QDate(2010, 1, 29), 94.68, 548623.75, 5794.51, "pink"},
     {"trait 22", 30., QDate(2010, 1, 30), 71.55, 413429.91, 5778.2, "pink"},
     {"trait 1", 31., QDate(2010, 1, 31), 80.52, 324537.82, 4030.52, "blue"},
     {"trait 2", 32., QDate(2010, 2, 1), 92.71, 473075.07, 5102.74, "black"},
     {"trait 3", 33., QDate(2010, 2, 2), 58.86, 339724.62, 5771.74, "blue"},
     {"trait 4", 34., QDate(2010, 2, 3), 70.17, 443233.11, 6316.56, "pink"},
     {"trait 5", 35., QDate(2010, 2, 4), 82.71, 426391.14, 5155.25, "black"},
     {"trait 6", 36., QDate(2010, 2, 5), 60.89, 425907.85, 6994.71, "blue"},
     {"trait 7", 37., QDate(2010, 2, 6), 64.22, 444124.57, 6915.67, "blue"},
     {"trait 8", 38., QDate(2010, 2, 7), 68.32, 353874.01, 5179.65, "brown"}},
    {},
    {{200000., 51., 3921.56862745098, QDate(2012, 02, 02), "a"},
     {200001., 52., 3846.17307692308, QDate(), "b"},
     {200002., 53., 3773.62264150943, QDate(2012, 02, 04), "b"},
     {200003., 54., 3703.75925925926, QDate(2012, 02, 05), "c"}},
    {{QVariant(QVariant::String), QDate(1970, 1, 1), 52.21, 1.47,
      QDate(1970, 1, 1), 0., "Sx", "105", 24.76, 0., 2725.8841, 0},
     {QVariant(QVariant::String), QDate(1970, 1, 2), 53.12, 1.5,
      QVariant(QVariant::Date), 1., "Sy", "931.17", 931.17, 1., 2821.7344,
      53.12},
     {QVariant(QVariant::String), QDate(1970, 1, 3), 54.48, 1.52,
      QVariant(QVariant::Date), 2., "Sxx", "1015", 41.0532, 4., 2968.0704,
      108.96},
     {QVariant(QVariant::String), QDate(1970, 1, 4), 55.84, 1.55,
      QVariant(QVariant::Date), 3., "Syy", "58498.5439", 58498.5439, 9.,
      3118.1056, 167.52},
     {QVariant(QVariant::String), QDate(1970, 1, 5), 57.2, 1.57,
      QVariant(QVariant::Date), 4., "Sxy", "6956.82", 1548.2453, 16., 3271.84,
      228.8},
     {QVariant(QVariant::String), QDate(1970, 1, 6), 58.57, 1.6,
      QVariant(QVariant::Date), 5., "n", "15", 15, 25., 3430.4449, 292.85},
     {QVariant(QVariant::String), QDate(1970, 1, 7), 59.93, 1.63,
      QVariant(QVariant::Date), 6., "a (beta w wiki)", "1.56653571428571",
      61.2721865421088, 36., 3591.6049, 359.58},
     {QVariant(QVariant::String), QDate(1970, 1, 8), 61.29, 1.65,
      QVariant(QVariant::Date), 7., "b (alfa w wiki)", "51.11225",
      -39.0619559188409, 49., 3756.4641, 429.03},
     {QVariant(QVariant::String), QDate(1970, 1, 9), 63.11, 1.68,
      QVariant(QVariant::Date), 8., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 64., 3982.8721, 504.88},
     {QVariant(QVariant::String), QDate(1970, 1, 10), 64.47, 1.7,
      QVariant(QVariant::Date), 9., QVariant(QVariant::String), "y=ax+b", 0.,
      81., 4156.3809, 580.23},
     {QVariant(QVariant::String), QDate(1970, 1, 11), 66.28, 1.73,
      QVariant(QVariant::Date), 10., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 100., 4393.0384, 662.8},
     {QVariant(QVariant::String), QDate(1970, 1, 12), 68.1, 1.75,
      QVariant(QVariant::Date), 11., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 121., 4637.61, 749.1},
     {QVariant(QVariant::String), QDate(1970, 1, 13), 69.92, 1.78,
      QVariant(QVariant::Date), 12., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 144., 4888.8064, 839.04},
     {QVariant(QVariant::String), QDate(1970, 1, 14), 72.19, 1.8,
      QVariant(QVariant::Date), 13., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 169., 5211.3961, 938.47},
     {QVariant(QVariant::String), QDate(1970, 1, 15), 74.46, 1.83,
      QVariant(QVariant::Date), 14., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 196., 5544.2916, 1042.44},
     {QVariant(QVariant::String), QVariant(QVariant::Date), 0., 0.,
      QVariant(QVariant::Date), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 2.1609, 2725.8841, 76.7487},
     {QVariant(QVariant::String), QVariant(QVariant::Date), 0., 0.,
      QVariant(QVariant::Date), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 2.25, 2821.7344, 79.68},
     {QVariant(QVariant::String), QVariant(QVariant::Date), 0., 0.,
      QVariant(QVariant::Date), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 2.3104, 2968.0704, 82.8096},
     {QVariant(QVariant::String), QVariant(QVariant::Date), 0., 0.,
      QVariant(QVariant::Date), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 2.4025, 3118.1056, 86.552},
     {QVariant(QVariant::String), QVariant(QVariant::Date), 0., 0.,
      QVariant(QVariant::Date), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 2.4649, 3271.84, 89.804},
     {QVariant(QVariant::String), QVariant(QVariant::Date), 0., 0.,
      QVariant(QVariant::Date), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 2.56, 3430.4449, 93.712},
     {QVariant(QVariant::String), QVariant(QVariant::Date), 0., 0.,
      QVariant(QVariant::Date), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 2.6569, 3591.6049, 97.6859},
     {QVariant(QVariant::String), QVariant(QVariant::Date), 0., 0.,
      QVariant(QVariant::Date), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 2.7225, 3756.4641, 101.1285},
     {QVariant(QVariant::String), QVariant(QVariant::Date), 0., 0.,
      QVariant(QVariant::Date), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 2.8224, 3982.8721, 106.0248},
     {QVariant(QVariant::String), QVariant(QVariant::Date), 0., 0.,
      QVariant(QVariant::Date), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 2.89, 4156.3809, 109.599},
     {QVariant(QVariant::String), QVariant(QVariant::Date), 0., 0.,
      QVariant(QVariant::Date), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 2.9929, 4393.0384, 114.6644},
     {QVariant(QVariant::String), QVariant(QVariant::Date), 0., 0.,
      QVariant(QVariant::Date), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 3.0625, 4637.61, 119.175},
     {QVariant(QVariant::String), QVariant(QVariant::Date), 0., 0.,
      QVariant(QVariant::Date), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 3.1684, 4888.8064, 124.4576},
     {QVariant(QVariant::String), QVariant(QVariant::Date), 0., 0.,
      QVariant(QVariant::Date), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 3.24, 5211.3961, 129.942},
     {QVariant(QVariant::String), QVariant(QVariant::Date), 0., 0.,
      QVariant(QVariant::Date), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 3.3489, 5544.2916, 136.2618}},
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
    {{56., 423000., 7553.57142857143, QVariant(QVariant::String),
      "Średnia:", 6715.2873295802},
     {45., 345000., 7666.66666666667, QVariant(QVariant::String),
      "Mediana:", 6277.77777777778},
     {76., 455664., 5995.57894736842, QVariant(QVariant::String), "Q10",
      4301.1655011655},
     {66., 330000., 5000., QVariant(QVariant::String), "Q25", 5065.78947368422},
     {46., 450000., 9782.60869565217, QVariant(QVariant::String), "Q50",
      6277.77777777778},
     {65., 280000., 4307.69230769231, QVariant(QVariant::String), "Q75",
      7668.62777777778},
     {65., 250000., 3846.15384615385, QVariant(QVariant::String), "Q90",
      9856.97940503433},
     {60., 380000., 6333.33333333333, QVariant(QVariant::String),
      QVariant(QVariant::String), 0.},
     {55., 450000., 8181.81818181818, QVariant(QVariant::String),
      QVariant(QVariant::String), 0.},
     {76., 800000., 10526.3157894737, QVariant(QVariant::String),
      QVariant(QVariant::String), 0.},
     {57., 290000., 5087.71929824562, QVariant(QVariant::String),
      QVariant(QVariant::String), 0.},
     {67., 410000., 6119.40298507463, QVariant(QVariant::String),
      QVariant(QVariant::String), 0.},
     {41., 310000., 7560.9756097561, QVariant(QVariant::String),
      QVariant(QVariant::String), 0.},
     {68., 390000., 5735.29411764706, QVariant(QVariant::String),
      QVariant(QVariant::String), 0.},
     {59., 400000., 6779.66101694915, QVariant(QVariant::String),
      QVariant(QVariant::String), 0.},
     {66., 280000., 4242.42424242424, QVariant(QVariant::String),
      QVariant(QVariant::String), 0.},
     {45., 280000., 6222.22222222222, QVariant(QVariant::String),
      QVariant(QVariant::String), 0.},
     {55., 270000., 4909.09090909091, QVariant(QVariant::String),
      QVariant(QVariant::String), 0.},
     {45., 345353., 7674.51111111111, QVariant(QVariant::String),
      QVariant(QVariant::String), 0.},
     {34., 366544., 10780.7058823529, QVariant(QVariant::String),
      QVariant(QVariant::String), 0.}},
    {{QDate(2012, 1, 1), "a", 45., QVariant(QVariant::String)},
     {QVariant(QVariant::Date), "s", 5., QVariant(QVariant::String)},
     {QVariant(QVariant::Date), "d", 67., QVariant(QVariant::String)}},
    {{"test1", "test123", "zawsze aktualna", QVariant(QVariant::String)},
     {"test2", "test123", QVariant(QVariant::String), "konto zablokowane"},
     {"test3", "test123", "2011-09-30", QVariant(QVariant::String)}}};

void ImportCommon::checkRetrievingSheetNames(ImportSpreadsheet& importer)
{
    auto [success, actualSheetNames] = importer.getSheetNames();
    QCOMPARE(success, true);
    QCOMPARE(actualSheetNames, sheetNames_);
}

void ImportCommon::checkRetrievingSheetNamesFromEmptyFile(
    ImportSpreadsheet& importer)
{
    auto [success, actualNames] = importer.getSheetNames();
    QCOMPARE(success, false);
    QCOMPARE(actualNames, {});
}

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

void ImportCommon::checkGetColumnList(ImportSpreadsheet& importer)
{
    QFETCH(QString, sheetName);
    QFETCH(QStringList, expectedColumnList);
    auto [success, actualColumnList] = importer.getColumnNames(sheetName);
    QCOMPARE(success, true);
    QCOMPARE(actualColumnList, expectedColumnList);
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

void ImportCommon::checkGetColumnTypes(ImportSpreadsheet& importer)
{
    QFETCH(QString, sheetName);
    QFETCH(QVector<ColumnType>, expectedColumnTypes);
    auto [success, columnTypes] = importer.getColumnTypes(sheetName);
    QCOMPARE(success, true);
    QCOMPARE(columnTypes, expectedColumnTypes);
}

void ImportCommon::checkSettingEmptyColumnName(ImportSpreadsheet& importer)
{
    const QString newEmptyColumnName{"<empty column>"};
    importer.setNameForEmptyColumn(newEmptyColumnName);
    auto [success, actualColumnList] = importer.getColumnNames(sheetNames_[4]);

    std::list<QString> expectedColumnList(
        static_cast<size_t>(testColumnNames_[4].size()));
    std::replace_copy(testColumnNames_[4].begin(), testColumnNames_[4].end(),
                      expectedColumnList.begin(), QString("---"),
                      newEmptyColumnName);

    QCOMPARE(success, true);
    QCOMPARE(actualColumnList, QStringList::fromStdList(expectedColumnList));
}

void ImportCommon::checkGetColumnListTwoSheets(ImportSpreadsheet& importer)
{
    auto [success, actualColumnList] = importer.getColumnNames(sheetNames_[4]);
    QCOMPARE(success, true);
    QCOMPARE(actualColumnList, testColumnNames_[4]);

    std::tie(success, actualColumnList) =
        importer.getColumnNames(sheetNames_[0]);
    QCOMPARE(success, true);
    QCOMPARE(actualColumnList, testColumnNames_[0]);
}

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

void ImportCommon::checkGetColumnCount(ImportSpreadsheet& importer)
{
    QFETCH(QString, sheetName);
    QFETCH(int, expectedColumnCount);
    auto [success, actualColumnCount] = importer.getColumnCount(sheetName);
    QCOMPARE(success, true);
    QCOMPARE(actualColumnCount, expectedColumnCount);
}

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

void ImportCommon::checkGetRowCount(ImportSpreadsheet& importer)
{
    QFETCH(QString, sheetName);
    QFETCH(unsigned int, expectedRowCount);
    auto [success, actualRowCount] = importer.getRowCount(sheetName);
    QCOMPARE(success, true);
    QCOMPARE(actualRowCount, expectedRowCount);
}

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

void ImportCommon::testGetRowAndColumnCountViaGetColumnTypes(
    ImportSpreadsheet& importer)
{
    QFETCH(QString, sheetName);
    QFETCH(unsigned int, expectedRowCount);
    QFETCH(unsigned int, expectedColumnCount);
    importer.getColumnTypes(sheetName);
    auto [successRowCount, actualRowCount] = importer.getRowCount(sheetName);
    QCOMPARE(successRowCount, true);
    QCOMPARE(actualRowCount, expectedRowCount);
    auto [successColumnCount, actualColumnCount] =
        importer.getColumnCount(sheetName);
    QCOMPARE(successColumnCount, true);
    QCOMPARE(actualColumnCount, expectedColumnCount);
}

void ImportCommon::prepareDataForGetData()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<QVector<QVector<QVariant>>>("expectedData");
    for (int i = 0; i < sheetData_.size(); ++i)
    {
        const QString& sheetName{sheetNames_[i]};
        QTest::newRow(("Data in " + sheetName).toStdString().c_str())
            << sheetName << sheetData_[i];
    }
}

void ImportCommon::checkGetData(ImportSpreadsheet& importer)
{
    QFETCH(QString, sheetName);
    QFETCH(QVector<QVector<QVariant>>, expectedData);
    auto [success, actualData] = importer.getData(sheetName, {});
    QCOMPARE(success, true);
    QCOMPARE(actualData, expectedData);
}

void ImportCommon::prepareDataForGetDataLimitRows()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<unsigned int>("rowLimit");
    QTest::addColumn<QVector<QVector<QVariant>>>("expectedData");
    QString sheetName{sheetNames_[0]};
    QTest::newRow(("Limited data to 10 in " + sheetName).toStdString().c_str())
        << sheetName << 10u << sheetData_[0];
    QTest::newRow(("Limited data to 2 in " + sheetName).toStdString().c_str())
        << sheetName << 2u << sheetData_[0];
    sheetName = sheetNames_[5];
    QVector<QVector<QVariant>> expectedValues;
    const unsigned int rowLimit{12u};
    for (unsigned int i = 0; i < rowLimit; ++i)
        expectedValues.append(sheetData_[5][static_cast<int>(i)]);
    QTest::newRow(
        ("Limited data to " + QString::number(rowLimit) + " in " + sheetName)
            .toStdString()
            .c_str())
        << sheetName << rowLimit << expectedValues;
}

void ImportCommon::checkGetDataLimitRows(ImportSpreadsheet& importer)
{
    QFETCH(QString, sheetName);
    QFETCH(unsigned int, rowLimit);
    QFETCH(QVector<QVector<QVariant>>, expectedData);
    auto [success, actualData] =
        importer.getLimitedData(sheetName, {}, rowLimit);
    QCOMPARE(success, true);
    QCOMPARE(actualData, expectedData);
}

void ImportCommon::prepareDataForGetGetDataExcludeColumns()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<QVector<unsigned int>>("excludedColumns");
    QTest::addColumn<QVector<QVector<QVariant>>>("expectedData");

    QString sheetName{sheetNames_[0]};
    QString testName{"Get data with excluded column 1 in " + sheetName};
    QTest::newRow(qUtf8Printable(testName))
        << sheetName << QVector<unsigned int>{1}
        << getDataWithoutColumns(sheetData_[0], {1});

    testName = "Get data with excluded column 0 and 2 in " + sheetName;
    QTest::newRow(qUtf8Printable(testName))
        << sheetName << QVector<unsigned int>{0, 2}
        << getDataWithoutColumns(sheetData_[0], {2, 0});
}

void ImportCommon::checkGetDataExcludeColumns(ImportSpreadsheet& importer)
{
    QFETCH(QString, sheetName);
    QFETCH(QVector<unsigned int>, excludedColumns);
    QFETCH(QVector<QVector<QVariant>>, expectedData);
    auto [success, actualData] = importer.getData(sheetName, excludedColumns);
    QCOMPARE(success, true);
    QCOMPARE(actualData, expectedData);
}

void ImportCommon::checkGetDataExcludeInvalidColumn(ImportSpreadsheet& importer)
{
    auto [success, actualData] = importer.getData(sheetNames_[0], {3});
    QCOMPARE(success, false);
}

QVector<QVector<QVariant>> ImportCommon::getDataForSheet(
    const QString& fileName)
{
    return sheetData_[sheetNames_.indexOf(fileName)];
}

QVector<QVector<QVariant>> ImportCommon::getDataWithoutColumns(
    const QVector<QVector<QVariant>>& data, QVector<int> columnsToExclude)
{
    QVector<QVector<QVariant>> expectedValues;
    for (auto dataRow : data)
    {
        for (int column : columnsToExclude)
            dataRow.remove(column);
        expectedValues.append(dataRow);
    }
    return expectedValues;
}

QStringList ImportCommon::getSheetNames() { return sheetNames_; }

void ImportCommon::checkInvalidSheetName(ImportSpreadsheet& importer)
{
    auto [success, actualSharedStrings] =
        importer.getColumnCount("invalidSheetName");
    QCOMPARE(success, false);
}

void ImportCommon::checkEmittingProgressPercentChangedEmptyFile(
    ImportSpreadsheet& importer)
{
    QSignalSpy spy(&importer, &ImportSpreadsheet::progressPercentChanged);
    QStringList sheetNames{importer.getSheetNames().second};
    auto [success, actualData] = importer.getData(sheetNames.first(), {});
    QCOMPARE(success, true);
    QCOMPARE(spy.count(), NO_SIGNAL);
}

void ImportCommon::checkEmittingProgressPercentChangedSmallFile(
    ImportSpreadsheet& importer)
{
    QSignalSpy spy(&importer, &ImportSpreadsheet::progressPercentChanged);
    const unsigned int sheetIndex{1};
    auto [success, actualData] = importer.getData(sheetNames_[sheetIndex], {});
    QCOMPARE(success, true);
    QCOMPARE(spy.count(), 19);
}

void ImportCommon::checkEmittingProgressPercentChangedBigFile(
    ImportSpreadsheet& importer)
{
    QSignalSpy spy(&importer, &ImportSpreadsheet::progressPercentChanged);
    QStringList sheetNames{importer.getSheetNames().second};
    auto [success, actualData] = importer.getData(sheetNames.first(), {});
    QCOMPARE(success, true);
    QCOMPARE(spy.count(), 100);
}
