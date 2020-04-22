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

    std::pair<bool, QStringList> getSheetList() override;
    std::pair<bool, QMap<QString, QString>> getSheetMap();
    void setSheetMap(QMap<QString, QString> sheetMap);

    std::pair<bool, QVector<ColumnType>> getColumnTypes(
        const QString& sheetName, int columnsCount) override;

    std::pair<bool, QStringList> getColumnList(
        const QString& sheetName) override;
    void setColumnList(QStringList columnList);

    std::pair<bool, QStringList> getSharedStrings();
    void setSharedStrings(QStringList sharedStrings);

    std::pair<bool, QList<int>> getDateStyles();
    void setDateStyles(QList<int> dateStyles);

    std::pair<bool, QList<int>> getAllStyles();
    void setAllStyles(QList<int> allStyles);

private:
    std::tuple<bool, std::optional<QList<int>>, std::optional<QList<int>>>
    getStyles();

    bool openZipAndMoveToSecondRow(QuaZip& zip, const QString& sheetName,
                                   QuaZipFile& zipFile,
                                   QXmlStreamReader& xmlStreamReader);

    std::optional<QMap<QString, QString>> sheetMap_{std::nullopt};

    std::optional<QStringList> columnList_{std::nullopt};

    std::optional<QStringList> sharedStrings_{std::nullopt};

    std::optional<QList<int>> dateStyles_{std::nullopt};

    std::optional<QList<int>> allStyles_{std::nullopt};

    static constexpr int NOT_SET_COLUMN{-1};

    static constexpr int DECIMAL_BASE{10};
};

#endif  // IMPORTXLSX_H
