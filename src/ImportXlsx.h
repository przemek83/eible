#ifndef IMPORTXLSX_H
#define IMPORTXLSX_H

#include <QHash>

#include "ImportSpreadsheet.h"

#include "eible_global.h"

class QuaZip;
class QuaZipFile;
class QXmlStreamReader;

class EIBLE_EXPORT ImportXlsx : public ImportSpreadsheet
{
public:
    explicit ImportXlsx(QIODevice& ioDevice);

    std::pair<bool, QStringList> getSheetNames() override;
    std::pair<bool, QList<std::pair<QString, QString>>> getSheets();
    void setSheets(QList<std::pair<QString, QString>> sheets);

    std::pair<bool, QVector<ColumnType>> getColumnTypes(
        const QString& sheetName) override;

    std::pair<bool, QStringList> getColumnNames(
        const QString& sheetName) override;

    std::pair<bool, QStringList> getSharedStrings();
    void setSharedStrings(QStringList sharedStrings);

    std::pair<bool, QList<int>> getDateStyles();
    void setDateStyles(QList<int> dateStyles);

    std::pair<bool, QList<int>> getAllStyles();
    void setAllStyles(QList<int> allStyles);

    std::pair<bool, unsigned int> getColumnCount(
        const QString& sheetName) override;

    std::pair<bool, unsigned int> getRowCount(
        const QString& sheetName) override;

    std::pair<bool, QVector<QVector<QVariant>>> getData(
        const QString& sheetName,
        const QVector<unsigned int>& excludedColumns) override;

    std::pair<bool, QVector<QVector<QVariant>>> getLimitedData(
        const QString& sheetName, const QVector<unsigned int>& excludedColumns,
        unsigned int rowLimit) override;

private:
    std::tuple<bool, std::optional<QList<int>>, std::optional<QList<int>>>
    getStyles();

    std::pair<bool, QString> getSheetPath(QString sheetName);

    bool openZipAndMoveToSecondRow(QuaZip& zip, const QString& sheetName,
                                   QuaZipFile& zipFile,
                                   QXmlStreamReader& xmlStreamReader);

    std::pair<bool, unsigned int> getCount(
        const QString& sheetName, const QHash<QString, unsigned int>& countMap);

    bool analyzeSheet(const QString& sheetName);

    std::optional<QList<std::pair<QString, QString>>> sheets_{std::nullopt};
    std::optional<QStringList> sharedStrings_{std::nullopt};
    std::optional<QList<int>> dateStyles_{std::nullopt};
    std::optional<QList<int>> allStyles_{std::nullopt};
    QHash<QString, unsigned int> rowCounts_{};
    QHash<QString, unsigned int> columnCounts_{};
    QHash<QString, QVector<ColumnType>> columnTypes_{};

    static constexpr int NOT_SET_COLUMN{-1};
    static constexpr int DECIMAL_BASE{10};
};

#endif  // IMPORTXLSX_H
