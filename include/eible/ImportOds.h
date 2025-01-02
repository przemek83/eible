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
        const QString& sheetName, const QVector<int>& excludedColumns,
        int rowLimit) override;

private:
    std::pair<bool, QStringList> getSheetNamesFromZipFile();

    std::pair<bool, int> getCount(const QString& sheetName,
                                  const QHash<QString, int>& countMap) override;

    std::pair<bool, QStringList> retrieveColumnNames(
        const QString& sheetName) override;

    std::tuple<bool, int, QVector<ColumnType>> retrieveRowCountAndColumnTypes(
        const QString& sheetName) override;

    bool moveToSecondRow(const QString& sheetName, QuaZipFile& quaZipFile,
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

    const QStringList RECOCNIZED_COLUMN_TYPES{STRING_TAG,   DATE_TAG,
                                              FLOAT_TAG,    PERCENTAGE_TAG,
                                              CURRENCY_TAG, TIME_TAG};

    const int ODS_STRING_DATE_LENGTH{10};
};
