#ifndef IMPORTODS_H
#define IMPORTODS_H

#include "ImportSpreadsheet.h"

#include <QHash>

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

    std::pair<bool, unsigned int> getColumnCount(
        const QString& sheetName) override;

    std::pair<bool, unsigned int> getRowCount(
        const QString& sheetName) override;

    std::pair<bool, QVector<QVector<QVariant>>> getLimitedData(
        const QString& sheetName, const QVector<unsigned int>& excludedColumns,
        unsigned int rowLimit) override;

private:
    std::pair<bool, unsigned int> getCount(
        const QString& sheetName, const QHash<QString, unsigned int>& countMap);

    std::pair<bool, QStringList> retrieveColumnNames(
        const QString& sheetName) override;

    std::tuple<bool, unsigned int, QVector<ColumnType>>
    retrieveRowCountAndColumnTypes(const QString& sheetName) override;

    bool moveToSecondRow(const QString& sheetName, QuaZipFile& zipFile,
                         QXmlStreamReader& xmlStreamReader);

    void skipToSheet(QXmlStreamReader& xmlStreamReader,
                     const QString& sheetName) const;

    bool isRecognizedColumnType(const QXmlStreamAttributes& attributes) const;

    int getColumnRepeatCount(const QXmlStreamAttributes& attributes) const;

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

    static const QString TABLE_TAG;
    static const QString TABLE_ROW_TAG;
    static const QString TABLE_CELL_TAG;
    static const QString OFFICE_VALUE_TYPE_TAG;
    static const QString COLUMNS_REPEATED_TAG;
    static const QString STRING_TAG;
    static const QString DATE_TAG;
    static const QString FLOAT_TAG;
    static const QString PERCENTAGE_TAG;
    static const QString CURRENCY_TAG;
    static const QString TIME_TAG;
    static const QString P_TAG;
    static const QString OFFICE_DATE_VALUE_TAG;
    static const QString OFFICE_VALUE_TAG;
    static const QString DATE_FORMAT;
    static const QString TABLE_QUALIFIED_NAME;
    static const QString TABLE_NAME_TAG;
};

#endif  // IMPORTODS_H
