#ifndef IMPORTXLSX_H
#define IMPORTXLSX_H

#include <QMap>

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

    std::pair<bool, QStringList> getColumnList(
        const QString& sheetName) override;
    void setColumnList(QStringList columnList);

    std::pair<bool, QStringList> getSharedStrings();
    void setSharedStrings(QStringList sharedStrings);

    std::pair<bool, QList<int>> getDateStyles();
    void setDateStyles(QList<int> dateStyles);

    std::pair<bool, QList<int>> getAllStyles();
    void setAllStyles(QList<int> allStyles);

    std::pair<bool, uint32_t> getColumnCount(const QString& name) override;

private:
    std::tuple<bool, std::optional<QList<int>>, std::optional<QList<int>>>
    getStyles();

    std::pair<bool, QString> getSheetPath(QString sheetName);

    bool openZipAndMoveToSecondRow(QuaZip& zip, const QString& sheetName,
                                   QuaZipFile& zipFile,
                                   QXmlStreamReader& xmlStreamReader);

    std::optional<QList<std::pair<QString, QString>>> sheets_{std::nullopt};
    std::optional<QStringList> columnList_{std::nullopt};
    std::optional<QStringList> sharedStrings_{std::nullopt};
    std::optional<QList<int>> dateStyles_{std::nullopt};
    std::optional<QList<int>> allStyles_{std::nullopt};

    static constexpr int NOT_SET_COLUMN{-1};
    static constexpr int DECIMAL_BASE{10};
};

#endif  // IMPORTXLSX_H
