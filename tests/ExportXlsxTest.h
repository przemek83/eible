#pragma once

#include <QObject>

class QAbstractItemView;
class TestTableModel;
class QBuffer;

class ExportXlsxTest : public QObject
{
public:
    explicit ExportXlsxTest(QObject* parent = nullptr);

    Q_OBJECT
private Q_SLOTS:
    void testExportingEmptyTable() const;

    void testExportingHeadersOnly() const;

    void testExportingSimpleTable() const;

    void testExportingViewWithMultiSelection() const;

    void benchmark_data();

    void benchmark();

    void cleanupTestCase();

private:
    static QByteArray retrieveFileFromZip(QBuffer& exportedZip,
                                          const QString& fileName);

    void compareWorkSheets(QBuffer& exportedZip,
                           const QString& sheetData) const;

    static void exportZip(const QAbstractItemView& view, QBuffer& exportedZip);

    QString zipWorkSheetPath_{QStringLiteral("xl/worksheets/sheet1.xml")};

    QString tableSheetData_{
        QString::fromLatin1(R"(<sheetData>)"
                            R"(<row r="1" spans="1:1" x14ac:dyDescent="0.25">)"
                            R"(<c r="A1" t="str" s="6"><v>Text</v></c>)"
                            R"(<c r="B1" t="str" s="6"><v>Numeric</v></c>)"
                            R"(<c r="C1" t="str" s="6"><v>Date</v></c>)"
                            R"(</row>)"
                            R"(<row r="2" spans="1:1" x14ac:dyDescent="0.25">)"
                            R"(<c r="A2" t="str"><v>Item 0 0</v></c>)"
                            R"(<c r="B2" s="4"><v>1</v></c>)"
                            R"(<c r="C2" s="3"><v>43833</v></c>)"
                            R"(</row>)"
                            R"(<row r="3" spans="1:1" x14ac:dyDescent="0.25">)"
                            R"(<c r="A3" t="str"><v>Item 0 1</v></c>)"
                            R"(<c r="B3" s="4"><v>2</v></c>)"
                            R"(<c r="C3" s="3"><v>43834</v></c>)"
                            R"(</row>)"
                            R"(<row r="4" spans="1:1" x14ac:dyDescent="0.25">)"
                            R"(<c r="A4" t="str"><v>Item 0 2</v></c>)"
                            R"(<c r="B4" s="4"><v>3</v></c>)"
                            R"(<c r="C4" s="3"><v>43835</v></c>)"
                            R"(</row>)"
                            R"(</sheetData>)")};

    QString multiSelectionTableSheetData_{
        QString::fromLatin1(R"(<sheetData>)"
                            R"(<row r="1" spans="1:1" x14ac:dyDescent="0.25">)"
                            R"(<c r="A1" t="str" s="6"><v>Text</v></c>)"
                            R"(<c r="B1" t="str" s="6"><v>Numeric</v></c>)"
                            R"(<c r="C1" t="str" s="6"><v>Date</v></c>)"
                            R"(</row>)"
                            R"(<row r="2" spans="1:1" x14ac:dyDescent="0.25">)"
                            R"(<c r="A2" t="str"><v>Item 0 0</v></c>)"
                            R"(<c r="B2" s="4"><v>1</v></c>)"
                            R"(<c r="C2" s="3"><v>43833</v></c>)"
                            R"(</row>)"
                            R"(<row r="3" spans="1:1" x14ac:dyDescent="0.25">)"
                            R"(<c r="A3" t="str"><v>Item 0 2</v></c>)"
                            R"(<c r="B3" s="4"><v>3</v></c>)"
                            R"(<c r="C3" s="3"><v>43835</v></c>)"
                            R"(</row>)"
                            R"(</sheetData>)")};

    QString headersOnlySheetData_{
        QString::fromLatin1(R"(<sheetData>)"
                            R"(<row r="1" spans="1:1" x14ac:dyDescent="0.25">)"
                            R"(<c r="A1" t="str" s="6"><v>Text</v></c>)"
                            R"(<c r="B1" t="str" s="6"><v>Numeric</v></c>)"
                            R"(<c r="C1" t="str" s="6"><v>Date</v></c>)"
                            R"(</row>)"
                            R"(</sheetData>)")};

    QString emptySheetData_{QStringLiteral(R"(</sheetData>)")};

    QStringList headers_{QStringLiteral("Text"), QStringLiteral("Numeric"),
                         QStringLiteral("Date")};

    TestTableModel* tableModelForBenchmarking_{nullptr};
};
