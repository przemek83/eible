#include "ImportXlsxTest.h"

#include <QBuffer>
#include <QSignalSpy>
#include <QTableView>
#include <QTest>

#include <ExportXlsx.h>
#include "ImportXlsx.h"

#include "TestTableModel.h"

const QList<std::pair<QString, QString>> ImportXlsxTest::sheets_{
    {"Sheet1", "xl/worksheets/sheet1.xml"},
    {"Sheet2", "xl/worksheets/sheet2.xml"},
    {"Sheet3(empty)", "xl/worksheets/sheet3.xml"},
    {"Sheet4", "xl/worksheets/sheet4.xml"},
    {"Sheet5", "xl/worksheets/sheet5.xml"},
    {"Sheet6", "xl/worksheets/sheet6.xml"},
    {"Sheet7", "xl/worksheets/sheet7.xml"},
    {"testAccounts", "xl/worksheets/sheet8.xml"}};

const QStringList ImportXlsxTest::sharedStrings_{"Text",
                                                 "Numeric",
                                                 "Date",
                                                 "Item 0, 0",
                                                 "Other item",
                                                 "Trait #1",
                                                 "Value #1",
                                                 "Transaction date",
                                                 "Units",
                                                 "Price",
                                                 "Price per unit",
                                                 "Another trait",
                                                 "trait 1",
                                                 "brown",
                                                 "trait 2",
                                                 "red",
                                                 "trait 3",
                                                 "yellow",
                                                 "trait 4",
                                                 "black",
                                                 "trait 16",
                                                 "trait 17",
                                                 "trait 18",
                                                 "white",
                                                 "trait 19",
                                                 "trait 20",
                                                 "trait 21",
                                                 "pink",
                                                 "trait 22",
                                                 "blue",
                                                 "trait 5",
                                                 "trait 6",
                                                 "trait 7",
                                                 "trait 8",
                                                 "cena nier",
                                                 "pow",
                                                 "cena metra",
                                                 "data transakcji",
                                                 "text",
                                                 "a",
                                                 "b",
                                                 "c",
                                                 "name",
                                                 "date",
                                                 "mass (kg)",
                                                 "height",
                                                 "Sx",
                                                 "Sy",
                                                 "Sxx",
                                                 "Syy",
                                                 "Sxy",
                                                 "n",
                                                 "a (beta w wiki)",
                                                 "b (alfa w wiki)",
                                                 "y=ax+b",
                                                 "modificator",
                                                 "x",
                                                 "y",
                                                 "Pow",
                                                 "Cena",
                                                 "cena_m",
                                                 "Średnia:",
                                                 "Mediana:",
                                                 "Q10",
                                                 "Q25",
                                                 "Q50",
                                                 "Q75",
                                                 "Q90",
                                                 "user",
                                                 "pass",
                                                 "lic_exp",
                                                 "uwagi",
                                                 "test1",
                                                 "test123",
                                                 "zawsze aktualna",
                                                 "test2",
                                                 "konto zablokowane",
                                                 "test3"};

const QList<int> ImportXlsxTest::dateStyles_{14,  15,  16,  17,  22,
                                             165, 167, 169, 170, 171};
const QList<int> ImportXlsxTest::allStyles_{164, 164, 165, 164, 166, 167, 168,
                                            168, 169, 164, 164, 170, 171};

const QList<QStringList> ImportXlsxTest::testColumnNames_ = {
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

const std::vector<unsigned int> ImportXlsxTest::expectedRowCounts_{
    2, 19, 0, 4, 30, 25, 20, 3};

const QVector<QVector<ColumnType>> ImportXlsxTest::columnTypes_ = {
    {ColumnType::STRING, ColumnType::NUMBER, ColumnType::DATE},
    {ColumnType::STRING, ColumnType::NUMBER, ColumnType::DATE,
     ColumnType::NUMBER, ColumnType::NUMBER, ColumnType::NUMBER,
     ColumnType::STRING},
    {},
    {ColumnType::NUMBER, ColumnType::NUMBER, ColumnType::NUMBER,
     ColumnType::DATE, ColumnType::STRING},
    {ColumnType::STRING, ColumnType::STRING, ColumnType::NUMBER,
     ColumnType::NUMBER, ColumnType::STRING, ColumnType::NUMBER,
     ColumnType::STRING, ColumnType::STRING, ColumnType::NUMBER,
     ColumnType::NUMBER, ColumnType::NUMBER, ColumnType::NUMBER},
    {ColumnType::NUMBER, ColumnType::NUMBER, ColumnType::NUMBER},
    {ColumnType::NUMBER, ColumnType::NUMBER, ColumnType::NUMBER,
     ColumnType::STRING, ColumnType::STRING, ColumnType::NUMBER},
    {ColumnType::STRING, ColumnType::STRING, ColumnType::STRING,
     ColumnType::STRING}};

const QVector<QVector<QVector<QVariant>>> ImportXlsxTest::sheetData_ = {
    {{3, 1., QDate(2020, 1, 3)}, {4, 2., QDate(2020, 1, 4)}},
    {{12, 1., QDate(2010, 1, 7), 48.49, 550060.8, 11343.8, 13},
     {14, 2., QDate(2010, 1, 8), 45.98, 236341.18, 5140.09, 15},
     {16, 9., QDate(2010, 1, 9), 51.33, 422955.71, 8239.93, 17},
     {18, 10., QDate(2010, 1, 10), 83.58, 372950.9, 4462.2, 19},
     {20, 22., QDate(2010, 1, 22), 33.06, 281118.11, 8503.27, 17},
     {21, 24., QDate(2010, 1, 24), 61.22, 352779.19, 5762.48, 19},
     {22, 25., QDate(2010, 1, 25), 81.16, 403985.68, 4977.65, 23},
     {24, 27., QDate(2010, 1, 27), 78.16, 410101.03, 5246.94, 19},
     {25, 28., QDate(2010, 1, 28), 73.66, 323727.84, 4394.89, 17},
     {26, 29., QDate(2010, 1, 29), 94.68, 548623.75, 5794.51, 27},
     {28, 30., QDate(2010, 1, 30), 71.55, 413429.91, 5778.2, 27},
     {12, 31., QDate(2010, 1, 31), 80.52, 324537.82, 4030.52, 29},
     {14, 32., QDate(2010, 2, 1), 92.71, 473075.07, 5102.74, 19},
     {16, 33., QDate(2010, 2, 2), 58.86, 339724.62, 5771.74, 29},
     {18, 34., QDate(2010, 2, 3), 70.17, 443233.11, 6316.56, 27},
     {30, 35., QDate(2010, 2, 4), 82.71, 426391.14, 5155.25, 19},
     {31, 36., QDate(2010, 2, 5), 60.89, 425907.85, 6994.71, 29},
     {32, 37., QDate(2010, 2, 6), 64.22, 444124.57, 6915.67, 29},
     {33, 38., QDate(2010, 2, 7), 68.32, 353874.01, 5179.65, 13}},
    {},
    {{200000., 51., 3921.56862745098, QDate(2012, 02, 02), 39},
     {200001., 52., 3846.17307692308, QDate(), 40},
     {200002., 53., 3773.62264150943, QDate(2012, 02, 04), 40},
     {200003., 54., 3703.75925925926, QDate(2012, 02, 05), 41}},
    {{QVariant(QVariant::String), "25569", 52.21, 1.47, "25569", 0., 46, "105",
      24.76, 0., 2725.8841, 0},
     {QVariant(QVariant::String), "25570", 53.12, 1.5,
      QVariant(QVariant::String), 1., 47, "931.17", 931.17, 1., 2821.7344,
      53.12},
     {QVariant(QVariant::String), "25571", 54.48, 1.52,
      QVariant(QVariant::String), 2., 48, "1015", 41.0532, 4., 2968.0704,
      108.96},
     {QVariant(QVariant::String), "25572", 55.84, 1.55,
      QVariant(QVariant::String), 3., 49, "58498.5439", 58498.5439, 9.,
      3118.1056, 167.52},
     {QVariant(QVariant::String), "25573", 57.2, 1.57,
      QVariant(QVariant::String), 4., 50, "6956.82", 1548.2453, 16., 3271.84,
      228.8},
     {QVariant(QVariant::String), "25574", 58.57, 1.6,
      QVariant(QVariant::String), 5., 51, "15", 15, 25., 3430.4449, 292.85},
     {QVariant(QVariant::String), "25575", 59.93, 1.63,
      QVariant(QVariant::String), 6., 52, "1.56653571428571", 61.2721865421088,
      36., 3591.6049, 359.58},
     {QVariant(QVariant::String), "25576", 61.29, 1.65,
      QVariant(QVariant::String), 7., 53, "51.11225", -39.0619559188409, 49.,
      3756.4641, 429.03},
     {QVariant(QVariant::String), "25577", 63.11, 1.68,
      QVariant(QVariant::String), 8., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 64., 3982.8721, 504.88},
     {QVariant(QVariant::String), "25578", 64.47, 1.7,
      QVariant(QVariant::String), 9., QVariant(QVariant::String), 54, 0., 81.,
      4156.3809, 580.23},
     {QVariant(QVariant::String), "25579", 66.28, 1.73,
      QVariant(QVariant::String), 10., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 100., 4393.0384, 662.8},
     {QVariant(QVariant::String), "25580", 68.1, 1.75,
      QVariant(QVariant::String), 11., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 121., 4637.61, 749.1},
     {QVariant(QVariant::String), "25581", 69.92, 1.78,
      QVariant(QVariant::String), 12., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 144., 4888.8064, 839.04},
     {QVariant(QVariant::String), "25582", 72.19, 1.8,
      QVariant(QVariant::String), 13., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 169., 5211.3961, 938.47},
     {QVariant(QVariant::String), "25583", 74.46, 1.83,
      QVariant(QVariant::String), 14., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 196., 5544.2916, 1042.44},
     {QVariant(QVariant::String), QVariant(QVariant::String), 0., 0.,
      QVariant(QVariant::String), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 2.1609, 2725.8841, 76.7487},
     {QVariant(QVariant::String), QVariant(QVariant::String), 0., 0.,
      QVariant(QVariant::String), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 2.25, 2821.7344, 79.68},
     {QVariant(QVariant::String), QVariant(QVariant::String), 0., 0.,
      QVariant(QVariant::String), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 2.3104, 2968.0704, 82.8096},
     {QVariant(QVariant::String), QVariant(QVariant::String), 0., 0.,
      QVariant(QVariant::String), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 2.4025, 3118.1056, 86.552},
     {QVariant(QVariant::String), QVariant(QVariant::String), 0., 0.,
      QVariant(QVariant::String), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 2.4649, 3271.84, 89.804},
     {QVariant(QVariant::String), QVariant(QVariant::String), 0., 0.,
      QVariant(QVariant::String), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 2.56, 3430.4449, 93.712},
     {QVariant(QVariant::String), QVariant(QVariant::String), 0., 0.,
      QVariant(QVariant::String), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 2.6569, 3591.6049, 97.6859},
     {QVariant(QVariant::String), QVariant(QVariant::String), 0., 0.,
      QVariant(QVariant::String), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 2.7225, 3756.4641, 101.1285},
     {QVariant(QVariant::String), QVariant(QVariant::String), 0., 0.,
      QVariant(QVariant::String), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 2.8224, 3982.8721, 106.0248},
     {QVariant(QVariant::String), QVariant(QVariant::String), 0., 0.,
      QVariant(QVariant::String), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 2.89, 4156.3809, 109.599},
     {QVariant(QVariant::String), QVariant(QVariant::String), 0., 0.,
      QVariant(QVariant::String), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 2.9929, 4393.0384, 114.6644},
     {QVariant(QVariant::String), QVariant(QVariant::String), 0., 0.,
      QVariant(QVariant::String), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 3.0625, 4637.61, 119.175},
     {QVariant(QVariant::String), QVariant(QVariant::String), 0., 0.,
      QVariant(QVariant::String), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 3.1684, 4888.8064, 124.4576},
     {QVariant(QVariant::String), QVariant(QVariant::String), 0., 0.,
      QVariant(QVariant::String), 0., QVariant(QVariant::String),
      QVariant(QVariant::String), 0., 3.24, 5211.3961, 129.942},
     {QVariant(QVariant::String), QVariant(QVariant::String), 0., 0.,
      QVariant(QVariant::String), 0., QVariant(QVariant::String),
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
    {{56., 423000., 7553.57142857143, QVariant(QVariant::String), 61,
      6715.2873295802},
     {45., 345000., 7666.66666666667, QVariant(QVariant::String), 62,
      6277.77777777778},
     {76., 455664., 5995.57894736842, QVariant(QVariant::String), 63,
      4301.1655011655},
     {66., 330000., 5000., QVariant(QVariant::String), 64, 5065.78947368422},
     {46., 450000., 9782.60869565217, QVariant(QVariant::String), 65,
      6277.77777777778},
     {65., 280000., 4307.69230769231, QVariant(QVariant::String), 66,
      7668.62777777778},
     {65., 250000., 3846.15384615385, QVariant(QVariant::String), 67,
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
    {{72, 73, 74, QVariant(QVariant::String)},
     {75, 73, QVariant(QVariant::String), 76},
     {77, 73, "40816", QVariant(QVariant::String)}}};

void ImportXlsxTest::testRetrievingSheetNames()
{
    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    auto [success, actualNames] = importXlsx.getSheetNames();
    QCOMPARE(success, true);
    QStringList sheets;
    for (const auto& [sheetName, sheetPath] : sheets_)
        sheets << sheetName;
    QCOMPARE(actualNames, sheets);
}

void ImportXlsxTest::testRetrievingSheetNamesFromEmptyFile()
{
    QByteArray byteArray;
    QBuffer emptyBuffer(&byteArray);
    ImportXlsx importXlsx(emptyBuffer);
    auto [success, actualNames] = importXlsx.getSheetNames();
    QCOMPARE(success, false);
    QCOMPARE(actualNames, {});
}

void ImportXlsxTest::testGetDateStyles()
{
    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    auto [success, actualDateStyle] = importXlsx.getDateStyles();
    QCOMPARE(success, true);
    QCOMPARE(actualDateStyle, dateStyles_);
}

void ImportXlsxTest::testGetDateStylesNoContent()
{
    QFile xlsxTestFile(QStringLiteral(":/template.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    auto [success, actualDateStyle] = importXlsx.getDateStyles();
    QCOMPARE(success, true);
    QCOMPARE(actualDateStyle, QList({14, 15, 16, 17, 22}));
}

void ImportXlsxTest::testGetAllStyles()
{
    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    auto [success, actualAllStyle] = importXlsx.getAllStyles();
    QCOMPARE(success, true);
    QCOMPARE(actualAllStyle, allStyles_);
}

void ImportXlsxTest::testGetAllStylesNoContent()
{
    QFile xlsxTestFile(QStringLiteral(":/template.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    auto [success, actualAllStyle] = importXlsx.getAllStyles();
    QCOMPARE(success, true);
    QCOMPARE(actualAllStyle, QList({0, 164, 10, 14, 4, 0, 0, 3}));
}

void ImportXlsxTest::testGetSharedStrings()
{
    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    auto [success, actualSharedStrings] = importXlsx.getSharedStrings();
    QCOMPARE(success, true);
    QCOMPARE(actualSharedStrings, sharedStrings_);
}

void ImportXlsxTest::testGetSharedStringsNoContent()
{
    QFile xlsxTestFile(QStringLiteral(":/template.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    auto [success, actualSharedStrings] = importXlsx.getSharedStrings();
    QCOMPARE(success, true);
    QCOMPARE(actualSharedStrings, {});
}

void ImportXlsxTest::testGetColumnList_data()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<QStringList>("expectedColumnList");
    for (int i = 0; i < testColumnNames_.size(); ++i)
    {
        const QString& sheetName{sheets_[i].first};
        QTest::newRow(("Columns in " + sheetName).toStdString().c_str())
            << sheetName << testColumnNames_[i];
    }
}

void ImportXlsxTest::testGetColumnList()
{
    QFETCH(QString, sheetName);
    QFETCH(QStringList, expectedColumnList);
    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    importXlsx.setSharedStrings(sharedStrings_);
    importXlsx.setSheets(sheets_);
    auto [success, actualColumnList] = importXlsx.getColumnNames(sheetName);
    QCOMPARE(success, true);
    QCOMPARE(actualColumnList, expectedColumnList);
}

void ImportXlsxTest::testSettingEmptyColumnName()
{
    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    const QString newEmptyColumnName{"<empty column>"};
    importXlsx.setNameForEmptyColumn(newEmptyColumnName);
    importXlsx.setSharedStrings(sharedStrings_);
    importXlsx.setSheets(sheets_);
    auto [success, actualColumnList] =
        importXlsx.getColumnNames(sheets_[4].first);

    std::list<QString> expectedColumnList(testColumnNames_[4].size());
    std::replace_copy(testColumnNames_[4].begin(), testColumnNames_[4].end(),
                      expectedColumnList.begin(), QString("---"),
                      newEmptyColumnName);

    QCOMPARE(success, true);
    QCOMPARE(actualColumnList, QStringList::fromStdList(expectedColumnList));
}

void ImportXlsxTest::testGetColumnListTwoSheets()
{
    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    importXlsx.setSharedStrings(sharedStrings_);
    importXlsx.setSheets(sheets_);
    auto [success, actualColumnList] =
        importXlsx.getColumnNames(sheets_[4].first);
    QCOMPARE(success, true);
    QCOMPARE(actualColumnList, testColumnNames_[4]);

    std::tie(success, actualColumnList) =
        importXlsx.getColumnNames(sheets_[0].first);
    QCOMPARE(success, true);
    QCOMPARE(actualColumnList, testColumnNames_[0]);
}

void ImportXlsxTest::testGetColumnTypes_data()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<QVector<ColumnType>>("expectedColumnTypes");

    for (int i = 0; i < testColumnNames_.size(); ++i)
    {
        const QString& sheetName{sheets_[i].first};
        QTest::newRow(("Columns types in " + sheetName).toStdString().c_str())
            << sheetName << columnTypes_[i];
    }
}

void ImportXlsxTest::testGetColumnTypes()
{
    QFETCH(QString, sheetName);
    QFETCH(QVector<ColumnType>, expectedColumnTypes);

    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    setCommonData(importXlsx);
    auto [success, columnTypes] = importXlsx.getColumnTypes(sheetName);
    QCOMPARE(success, true);
    QCOMPARE(columnTypes, expectedColumnTypes);
}

void ImportXlsxTest::testGetColumnCount_data()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<int>("expectedColumnCount");
    for (int i = 0; i < testColumnNames_.size(); ++i)
    {
        const QString& sheetName{sheets_[i].first};
        QTest::newRow(("Columns in " + sheetName).toStdString().c_str())
            << sheetName << testColumnNames_[i].size();
    }
}

void ImportXlsxTest::testGetColumnCount()
{
    QFETCH(QString, sheetName);
    QFETCH(int, expectedColumnCount);

    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    setCommonData(importXlsx);
    auto [success, actualColumnCount] = importXlsx.getColumnCount(sheetName);
    QCOMPARE(success, true);
    QCOMPARE(actualColumnCount, expectedColumnCount);
}

void ImportXlsxTest::testGetRowCount_data()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<unsigned int>("expectedRowCount");
    for (int i = 0; i < testColumnNames_.size(); ++i)
    {
        const QString& sheetName{sheets_[i].first};
        QTest::newRow(("Rows in " + sheetName).toStdString().c_str())
            << sheetName << expectedRowCounts_[i];
    }
}

void ImportXlsxTest::testGetRowCount()
{
    QFETCH(QString, sheetName);
    QFETCH(unsigned int, expectedRowCount);

    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    setCommonData(importXlsx);
    auto [success, actualRowCount] = importXlsx.getRowCount(sheetName);
    QCOMPARE(success, true);
    QCOMPARE(actualRowCount, expectedRowCount);
}

void ImportXlsxTest::testGetRowAndColumnCountViaGetColumnTypes_data()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<unsigned int>("expectedRowCount");
    QTest::addColumn<unsigned int>("expectedColumnCount");

    for (int i = 0; i < testColumnNames_.size(); ++i)
    {
        const QString& sheetName{sheets_[i].first};
        QTest::newRow(("Rows and columns via GetColumnTypes() in " + sheetName)
                          .toStdString()
                          .c_str())
            << sheetName << expectedRowCounts_[i]
            << static_cast<unsigned int>(testColumnNames_[i].size());
    }
}

void ImportXlsxTest::testGetRowAndColumnCountViaGetColumnTypes()
{
    QFETCH(QString, sheetName);
    QFETCH(unsigned int, expectedRowCount);
    QFETCH(unsigned int, expectedColumnCount);

    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    setCommonData(importXlsx);
    importXlsx.getColumnTypes(sheetName);
    auto [successRowCount, actualRowCount] = importXlsx.getRowCount(sheetName);
    QCOMPARE(successRowCount, true);
    QCOMPARE(actualRowCount, expectedRowCount);
    auto [successColumnCount, actualColumnCount] =
        importXlsx.getColumnCount(sheetName);
    QCOMPARE(successColumnCount, true);
    QCOMPARE(actualColumnCount, expectedColumnCount);
}

void ImportXlsxTest::testGetData_data()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<QVector<QVector<QVariant>>>("expectedData");
    for (int i = 0; i < testColumnNames_.size(); ++i)
    {
        const QString& sheetName{sheets_[i].first};
        QTest::newRow(("Data in " + sheetName).toStdString().c_str())
            << sheetName << sheetData_[i];
    }
}

void ImportXlsxTest::testGetData()
{
    QFETCH(QString, sheetName);
    QFETCH(QVector<QVector<QVariant>>, expectedData);

    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    setCommonData(importXlsx);
    auto [success, actualData] = importXlsx.getData(sheetName, {});
    QCOMPARE(success, true);
    QCOMPARE(actualData, expectedData);
}

void ImportXlsxTest::testGetDataLimitRows_data()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<unsigned int>("rowLimit");
    QTest::addColumn<QVector<QVector<QVariant>>>("expectedData");
    QString sheetName{sheets_[0].first};
    QTest::newRow(("Limited data to 10 in " + sheetName).toStdString().c_str())
        << sheetName << 10u << sheetData_[0];
    QTest::newRow(("Limited data to 2 in " + sheetName).toStdString().c_str())
        << sheetName << 2u << sheetData_[0];
    sheetName = sheets_[5].first;
    QVector<QVector<QVariant>> expectedValues;
    const unsigned int rowLimit{12u};
    for (unsigned int i = 0; i < rowLimit; ++i)
        expectedValues.append(sheetData_[5][i]);
    QTest::newRow(
        ("Limited data to " + QString::number(rowLimit) + " in " + sheetName)
            .toStdString()
            .c_str())
        << sheetName << rowLimit << expectedValues;
}

void ImportXlsxTest::testGetDataLimitRows()
{
    QFETCH(QString, sheetName);
    QFETCH(unsigned int, rowLimit);
    QFETCH(QVector<QVector<QVariant>>, expectedData);

    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    setCommonData(importXlsx);
    auto [success, actualData] =
        importXlsx.getLimitedData(sheetName, {}, rowLimit);
    QCOMPARE(success, true);
    QCOMPARE(actualData, expectedData);
}

void ImportXlsxTest::testGetDataExcludeColumns_data()
{
    QTest::addColumn<QString>("sheetName");
    QTest::addColumn<QVector<unsigned int>>("excludedColumns");
    QTest::addColumn<QVector<QVector<QVariant>>>("expectedData");

    QString sheetName{sheets_[0].first};
    QString testName{"Get data with excluded column 1 in " + sheetName};
    QTest::newRow(qUtf8Printable(testName))
        << sheetName << QVector<unsigned int>{1}
        << getDataWithoutColumns(sheetData_[0], {1});

    testName = "Get data with excluded column 0 and 2 in " + sheetName;
    QTest::newRow(qUtf8Printable(testName))
        << sheetName << QVector<unsigned int>{0, 2}
        << getDataWithoutColumns(sheetData_[0], {2, 0});
}

void ImportXlsxTest::testGetDataExcludeColumns()
{
    QFETCH(QString, sheetName);
    QFETCH(QVector<unsigned int>, excludedColumns);
    QFETCH(QVector<QVector<QVariant>>, expectedData);

    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    setCommonData(importXlsx);
    auto [success, actualData] = importXlsx.getData(sheetName, excludedColumns);
    QCOMPARE(success, true);
    QCOMPARE(actualData, expectedData);
}

void ImportXlsxTest::testGetDataExcludeInvalidColumn()
{
    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    setCommonData(importXlsx);
    auto [success, actualData] = importXlsx.getData(sheets_[0].first, {3});
    QCOMPARE(success, false);
}

void ImportXlsxTest::benchmarkGetData_data()
{
    QSKIP("Skip preparing benchmark data.");
    TestTableModel model(100, 100000);
    QTableView view;
    view.setModel(&model);

    generatedXlsx_.clear();
    QBuffer exportedFile(&generatedXlsx_);
    exportedFile.open(QIODevice::WriteOnly);
    ExportXlsx exportXlsx;
    exportXlsx.exportView(view, exportedFile);
}

void ImportXlsxTest::benchmarkGetData()
{
    QSKIP("Skip benchmark.");
    QBuffer testFile(&generatedXlsx_);
    ImportXlsx importXlsx(testFile);
    auto [success, sheetNames] = importXlsx.getSheetNames();
    QCOMPARE(success, true);
    QCOMPARE(sheetNames.size(), 1);
    QBENCHMARK { importXlsx.getData(sheetNames.front(), {}); }
}

void ImportXlsxTest::testEmittingProgressPercentChangedEmptyFile()
{
    QFile xlsxTestFile(QStringLiteral(":/template.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    QSignalSpy spy(&importXlsx, &ImportSpreadsheet::progressPercentChanged);
    QStringList sheetNames{importXlsx.getSheetNames().second};
    auto [success, actualData] = importXlsx.getData(sheetNames.first(), {});
    QCOMPARE(success, true);
    QCOMPARE(spy.count(), NO_SIGNAL);
}

void ImportXlsxTest::testEmittingProgressPercentChangedSmallFile()
{
    QFile xlsxTestFile(QStringLiteral(":/testXlsx.xlsx"));
    ImportXlsx importXlsx(xlsxTestFile);
    QSignalSpy spy(&importXlsx, &ImportSpreadsheet::progressPercentChanged);
    setCommonData(importXlsx);
    const unsigned int sheetIndex{1};
    auto [success, actualData] =
        importXlsx.getData(sheets_[sheetIndex].first, {});
    QCOMPARE(success, true);
    QCOMPARE(spy.count(), expectedRowCounts_[sheetIndex]);
}

void ImportXlsxTest::testEmittingProgressPercentChangedBigFile_data()
{
    TestTableModel model(1, 250);
    QTableView view;
    view.setModel(&model);

    generatedXlsx_.clear();
    QBuffer exportedFile(&generatedXlsx_);
    exportedFile.open(QIODevice::WriteOnly);
    ExportXlsx exportXlsx;
    exportXlsx.exportView(view, exportedFile);
}

void ImportXlsxTest::testEmittingProgressPercentChangedBigFile()
{
    QBuffer testFile(&generatedXlsx_);
    ImportXlsx importXlsx(testFile);
    QSignalSpy spy(&importXlsx, &ImportSpreadsheet::progressPercentChanged);
    QStringList sheetNames{importXlsx.getSheetNames().second};
    auto [success, actualData] = importXlsx.getData(sheetNames.first(), {});
    QCOMPARE(success, true);
    QCOMPARE(spy.count(), 100);
}

QVector<QVector<QVariant>> ImportXlsxTest::getDataWithoutColumns(
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

void ImportXlsxTest::setCommonData(ImportXlsx& importXlsx)
{
    importXlsx.setSharedStrings(sharedStrings_);
    importXlsx.setDateStyles(dateStyles_);
    importXlsx.setAllStyles(allStyles_);
    importXlsx.setSheets(sheets_);
}
