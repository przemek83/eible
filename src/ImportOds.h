#ifndef IMPORTODS_H
#define IMPORTODS_H

#include "ImportSpreadsheet.h"

#include "eible_global.h"

#include <QHash>

class QuaZip;
class QuaZipFile;
class QXmlStreamReader;

class EIBLE_EXPORT ImportOds : public ImportSpreadsheet
{
public:
    explicit ImportOds(QIODevice& ioDevice);

    std::pair<bool, QStringList> getSheetNames() override;
    //    std::pair<bool, QList<std::pair<QString, QString>>> getSheets();
    //    void setSheets(QList<std::pair<QString, QString>> sheets);

    std::pair<bool, QVector<ColumnType>> getColumnTypes(
        const QString& sheetName) override;

    std::pair<bool, QStringList> getColumnNames(
        const QString& sheetName) override;

    //    std::pair<bool, QStringList> getSharedStrings();
    //    void setSharedStrings(QStringList sharedStrings);

    //    std::pair<bool, QList<int>> getDateStyles();
    //    void setDateStyles(QList<int> dateStyles);

    //    std::pair<bool, QList<int>> getAllStyles();
    //    void setAllStyles(QList<int> allStyles);

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
    std::pair<bool, unsigned int> getCount(
        const QString& sheetName, const QHash<QString, unsigned int>& countMap);

    bool analyzeSheet(const QString& sheetName);

    bool openZipAndMoveToSecondRow(QuaZip& zip, const QString& sheetName,
                                   QuaZipFile& zipFile,
                                   QXmlStreamReader& xmlStreamReader);

    std::optional<QStringList> sheetNames_{std::nullopt};
    QHash<QString, unsigned int> rowCounts_{};
    QHash<QString, unsigned int> columnCounts_{};
};

#endif  // IMPORTODS_H
