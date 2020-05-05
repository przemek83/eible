#ifndef IMPORTODS_H
#define IMPORTODS_H

#include "ImportSpreadsheet.h"

#include <quazip5/quazip.h>
#include <QHash>

#include "eible_global.h"

class QuaZipFile;
class QXmlStreamReader;
class QXmlStreamAttributes;

class EIBLE_EXPORT ImportOds : public ImportSpreadsheet
{
public:
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

    bool analyzeSheet(const QString& sheetName);

    std::pair<unsigned int, unsigned int> getRowAndColumnCount(
        QXmlStreamReader& xmlStreamReader) const;

    bool moveToSecondRow(const QString& sheetName, QuaZipFile& zipFile,
                         QXmlStreamReader& xmlStreamReader);

    void skipToSheet(QXmlStreamReader& xmlStreamReader,
                     const QString& sheetName) const;

    bool openZipFile(QuaZipFile& zipFile, const QString& zipFileName);

    bool isRecognizedColumnType(const QXmlStreamAttributes& attributes) const;

    int getColumnRepeatCount(const QXmlStreamAttributes& attributes) const;

    bool isRowStart(const QXmlStreamReader& xmlStreamReader) const;
    bool isRowEnd(const QXmlStreamReader& xmlStreamReader) const;
    bool isCellStart(const QXmlStreamReader& xmlStreamReader) const;
    bool isCellEnd(const QXmlStreamReader& xmlStreamReader) const;

    std::optional<QStringList> sheetNames_{std::nullopt};
    QHash<QString, unsigned int> rowCounts_{};
    QHash<QString, unsigned int> columnCounts_{};
    QHash<QString, QVector<ColumnType>> columnTypes_{};
    QuaZip zip_;

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
