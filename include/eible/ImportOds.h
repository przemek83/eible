#pragma once

#include <optional>

#include <QHash>

#include "ImportSpreadsheet.h"
#include "eible_global.h"

class QuaZipFile;
class QXmlStreamReader;
class QXmlStreamAttributes;

/**
 * @class ImportOds
 * @brief Class analysing content and loading it from ods files.
 */
class EIBLE_EXPORT ImportOds : public ImportSpreadsheet
{
    Q_OBJECT
public:
    /**
     * @brief Constructor.
     * @param ioDevice Source of data (QFile, QBuffer, ...).
     */
    explicit ImportOds(QIODevice& ioDevice);

    std::pair<bool, QStringList> getSheetNames() override;

    std::pair<bool, QVector<ColumnType>> getColumnTypes(
        const QString& sheetName) override;

    std::pair<bool, QStringList> getColumnNames(
        const QString& sheetName) override;

    std::pair<bool, QVector<QVector<QVariant>>> getLimitedData(
        const QString& sheetName, const QVector<unsigned int>& excludedColumns,
        unsigned int rowLimit) override;

private:
    std::pair<bool, unsigned int> getCount(
        const QString& sheetName,
        const QHash<QString, unsigned int>& countMap) override;

    std::pair<bool, QStringList> retrieveColumnNames(
        const QString& sheetName) override;

    std::tuple<bool, unsigned int, QVector<ColumnType>>
    retrieveRowCountAndColumnTypes(const QString& sheetName) override;

    bool moveToSecondRow(const QString& sheetName, QuaZipFile& zipFile,
                         QXmlStreamReader& xmlStreamReader) const;

    void skipToSheet(QXmlStreamReader& xmlStreamReader,
                     const QString& sheetName) const;

    bool isRecognizedColumnType(const QXmlStreamAttributes& attributes) const;

    unsigned int getColumnRepeatCount(
        const QXmlStreamAttributes& attributes) const;

    bool isRowStart(const QXmlStreamReader& xmlStreamReader) const;
    bool isRowEnd(const QXmlStreamReader& xmlStreamReader) const;
    bool isCellStart(const QXmlStreamReader& xmlStreamReader) const;
    bool isCellEnd(const QXmlStreamReader& xmlStreamReader) const;

    ColumnType recognizeColumnType(ColumnType currentType,
                                   const QString& xmlColTypeValue) const;

    QVariant retrieveValueFromField(QXmlStreamReader& xmlStreamReader,
                                    ColumnType columnType) const;

    bool isOfficeValueTagEmpty(const QXmlStreamReader& xmlStreamReader) const;

    std::optional<QStringList> sheetNames_{std::nullopt};

    const QString TABLE_TAG{QStringLiteral("table")};
    const QString TABLE_ROW_TAG{QStringLiteral("table-row")};
    const QString TABLE_CELL_TAG{QStringLiteral("table-cell")};
    const QString OFFICE_VALUE_TYPE_TAG{QStringLiteral("office:value-type")};
    const QString COLUMNS_REPEATED_TAG{
        QStringLiteral("table:number-columns-repeated")};
    const QString STRING_TAG{QStringLiteral("string")};
    const QString DATE_TAG{QStringLiteral("date")};
    const QString FLOAT_TAG{QStringLiteral("float")};
    const QString PERCENTAGE_TAG{QStringLiteral("percentage")};
    const QString CURRENCY_TAG{QStringLiteral("currency")};
    const QString TIME_TAG{QStringLiteral("time")};
    const QString P_TAG{QStringLiteral("p")};
    const QString OFFICE_DATE_VALUE_TAG{QStringLiteral("office:date-value")};
    const QString OFFICE_VALUE_TAG{QStringLiteral("office:value")};
    const QString DATE_FORMAT{QStringLiteral("yyyy-MM-dd")};
    const QString TABLE_QUALIFIED_NAME{QStringLiteral("table:table")};
    const QString TABLE_NAME_TAG{QStringLiteral("table:name")};
};
