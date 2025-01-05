#pragma once

#include <QObject>
#include <QSet>
#include <QVector>

#include <eible/ColumnType.h>

class ImportXlsx;

class ImportXlsxTest : public QObject
{
public:
    explicit ImportXlsxTest(QObject* parent = nullptr);

    Q_OBJECT
private Q_SLOTS:
    void testRetrievingSheetNames() const;
    static void testRetrievingSheetNamesFromEmptyFile();

    void testGetDateStyles() const;
    void testGetDateStylesNoContent() const;

    void testGetAllStyles() const;
    void testGetAllStylesNoContent() const;

    void testGetSharedStrings() const;
    void testGetSharedStringsNoContent() const;

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

    void testGetData_data();
    void testGetData() const;

    void testGetDataLimitRows_data();
    void testGetDataLimitRows() const;

    void testGetDataExcludeColumns_data();
    void testGetDataExcludeColumns() const;
    void testGetDataExcludeInvalidColumn() const;

    void benchmarkGetData_data();
    void benchmarkGetData();

    void testEmittingProgressPercentChangedEmptyFile() const;
    void testEmittingProgressPercentChangedSmallFile() const;
    static void testEmittingProgressPercentChangedBigFile();

    void testInvalidSheetName() const;

    static void testDamagedFile();

private:
    QVector<QVector<QVariant>> convertDataToUseSharedStrings(
        const QVector<QVector<QVariant>>& inputData) const;

    const QString testFileName_{QStringLiteral(":/testXlsx.xlsx")};
    const QString templateFileName_{QStringLiteral(":/template.xlsx")};

    const QVector<std::pair<QString, QString>> sheets_{
        {"Sheet1", "xl/worksheets/sheet1.xml"},
        {"Sheet2", "xl/worksheets/sheet2.xml"},
        {"Sheet3(empty)", "xl/worksheets/sheet3.xml"},
        {"Sheet4", "xl/worksheets/sheet4.xml"},
        {"Sheet5", "xl/worksheets/sheet5.xml"},
        {"Sheet6", "xl/worksheets/sheet6.xml"},
        {"Sheet7", "xl/worksheets/sheet7.xml"},
        {"Sheet9", "xl/worksheets/sheet8.xml"},
        {"testAccounts", "xl/worksheets/sheet9.xml"}};

    const QStringList sharedStrings_{QStringLiteral("Text"),
                                     QStringLiteral("Numeric"),
                                     QStringLiteral("Date"),
                                     QStringLiteral("Item 0, 0"),
                                     QStringLiteral("Other item"),
                                     QStringLiteral("Trait #1"),
                                     QStringLiteral("Value #1"),
                                     QStringLiteral("Transaction date"),
                                     QStringLiteral("Units"),
                                     QStringLiteral("Price"),
                                     QStringLiteral("Price per unit"),
                                     QStringLiteral("Another trait"),
                                     QStringLiteral("trait 1"),
                                     QStringLiteral("brown"),
                                     QStringLiteral("trait 2"),
                                     QStringLiteral("red"),
                                     QStringLiteral("trait 3"),
                                     QStringLiteral("yellow"),
                                     QStringLiteral("trait 4"),
                                     QStringLiteral("black"),
                                     QStringLiteral("trait 16"),
                                     QStringLiteral("trait 17"),
                                     QStringLiteral("trait 18"),
                                     QStringLiteral("white"),
                                     QStringLiteral("trait 19"),
                                     QStringLiteral("trait 20"),
                                     QStringLiteral("trait 21"),
                                     QStringLiteral("pink"),
                                     QStringLiteral("trait 22"),
                                     QStringLiteral("blue"),
                                     QStringLiteral("trait 5"),
                                     QStringLiteral("trait 6"),
                                     QStringLiteral("trait 7"),
                                     QStringLiteral("trait 8"),
                                     QStringLiteral("cena nier"),
                                     QStringLiteral("pow"),
                                     QStringLiteral("cena metra"),
                                     QStringLiteral("data transakcji"),
                                     QStringLiteral("text"),
                                     QStringLiteral("a"),
                                     QStringLiteral("b"),
                                     QStringLiteral("c"),
                                     QStringLiteral("name"),
                                     QStringLiteral("date"),
                                     QStringLiteral("mass (kg)"),
                                     QStringLiteral("height"),
                                     QStringLiteral("Sx"),
                                     QStringLiteral("Sy"),
                                     QStringLiteral("Sxx"),
                                     QStringLiteral("Syy"),
                                     QStringLiteral("Sxy"),
                                     QStringLiteral("n"),
                                     QStringLiteral("a (beta w wiki)"),
                                     QStringLiteral("b (alfa w wiki)"),
                                     QStringLiteral("y=ax+b"),
                                     QStringLiteral("modificator"),
                                     QStringLiteral("x"),
                                     QStringLiteral("y"),
                                     QStringLiteral("Pow"),
                                     QStringLiteral("Cena"),
                                     QStringLiteral("cena_m"),
                                     QStringLiteral("Średnia:"),
                                     QStringLiteral("Mediana:"),
                                     QStringLiteral("Q10"),
                                     QStringLiteral("Q25"),
                                     QStringLiteral("Q50"),
                                     QStringLiteral("Q75"),
                                     QStringLiteral("Q90"),
                                     QStringLiteral("second"),
                                     QStringLiteral("third"),
                                     QStringLiteral("fourth"),
                                     QStringLiteral("s"),
                                     QStringLiteral("d"),
                                     QStringLiteral("user"),
                                     QStringLiteral("pass"),
                                     QStringLiteral("lic_exp"),
                                     QStringLiteral("uwagi"),
                                     QStringLiteral("test1"),
                                     QStringLiteral("test123"),
                                     QStringLiteral("zawsze aktualna"),
                                     QStringLiteral("test2"),
                                     QStringLiteral("konto zablokowane"),
                                     QStringLiteral("test3")};

    const QList<int> dateStyles_{14,  15,  16,  17,  22, 165,
                                 167, 170, 171, 172, 173};
    const QList<int> allStyles_{164, 164, 165, 164, 166, 167, 168, 169,
                                169, 170, 164, 164, 171, 172, 173};

    QByteArray generatedXlsx_;

    static constexpr int NO_SIGNAL{0};
};
