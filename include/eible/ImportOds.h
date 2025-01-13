#pragma once

#include <optional>

#include <QHash>

#include "ImportSpreadsheet.h"
#include "eible_global.h"

class QuaZipFile;
class QXmlStreamReader;
class QXmlStreamAttributes;

/// @class ImportOds
/// @brief Class analysing content and loading it from ods files.
class EIBLE_EXPORT ImportOds : public ImportSpreadsheet
{
    Q_OBJECT
public:
    /// @brief Constructor.
    /// @param ioDevice Source of data (QFile, QBuffer, ...).
    explicit ImportOds(QIODevice& ioDevice);

    std::pair<bool, QStringList> getSheetNames() override;

    std::pair<bool, QVector<ColumnType>> getColumnTypes(
        const QString& sheetName) override;

    std::pair<bool, QStringList> getColumnNames(
        const QString& sheetName) override;

    std::pair<bool, QVector<QVector<QVariant>>> getLimitedData(
        const QString& sheetName, const QVector<int>& excludedColumns,
        int rowLimit) override;

private:
    std::pair<bool, QStringList> getSheetNamesFromZipFile();

    std::pair<bool, int> getCount(const QString& sheetName,
                                  const QHash<QString, int>& countMap) override;

    std::pair<bool, QStringList> retrieveColumnNames(
        const QString& sheetName) override;

    void updateColumnTypes(QVector<ColumnType>& columnTypes, int column,
                           const QString& xmlColTypeValue, int repeats) const;

    std::tuple<bool, int, QVector<ColumnType>> retrieveRowCountAndColumnTypes(
        const QString& sheetName) override;

    void moveToSecondRow(const QString& sheetName, QuaZipFile& quaZipFile,
                         QXmlStreamReader& reader) const;

    void skipToSheet(QXmlStreamReader& reader, const QString& sheetName) const;

    bool isRecognizedColumnType(const QXmlStreamAttributes& attributes) const;

    int getColumnRepeatCount(const QXmlStreamAttributes& attributes) const;

    bool isRowStart(const QXmlStreamReader& reader) const;
    bool isRowEnd(const QXmlStreamReader& reader) const;
    bool isCellStart(const QXmlStreamReader& reader) const;
    bool isCellEnd(const QXmlStreamReader& reader) const;

    ColumnType recognizeColumnType(ColumnType currentType,
                                   const QString& xmlColTypeValue) const;

    QVariant retrieveValueFromStringColumn(QXmlStreamReader& reader) const;

    QVariant retrieveValueFromField(QXmlStreamReader& reader,
                                    ColumnType columnType) const;

    bool isOfficeValueTagEmpty(const QXmlStreamReader& reader) const;

    bool isSheetAvailable(const QString& sheetName);

    bool initializeColumnTypes(const QString& sheetName);

    std::optional<QStringList> sheetNames_{std::nullopt};

    const QString tableTag_{QStringLiteral("table")};
    const QString tableRowTag_{QStringLiteral("table-row")};
    const QString tableCellTag_{QStringLiteral("table-cell")};
    const QString officeValueTypeTag_{QStringLiteral("office:value-type")};
    const QString columnsRepeatedTag_{
        QStringLiteral("table:number-columns-repeated")};
    const QString stringTag_{QStringLiteral("string")};
    const QString dateTag_{QStringLiteral("date")};
    const QString floatTag_{QStringLiteral("float")};
    const QString percentageTag_{QStringLiteral("percentage")};
    const QString currencyTag_{QStringLiteral("currency")};
    const QString timeTag_{QStringLiteral("time")};
    const QString pTag_{QStringLiteral("p")};
    const QString officeDateValueTag_{QStringLiteral("office:date-value")};
    const QString officeValueTag_{QStringLiteral("office:value")};
    const QString dateFormat_{QStringLiteral("yyyy-MM-dd")};
    const QString tableQualifiedName_{QStringLiteral("table:table")};
    const QString tableNameTag_{QStringLiteral("table:name")};

    const QStringList recognizedColumnTypes_{stringTag_,   dateTag_,
                                             floatTag_,    percentageTag_,
                                             currencyTag_, timeTag_};

    const int odsStringDateLength_{10};
};
