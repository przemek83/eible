#ifndef IMPORTXLSX_H
#define IMPORTXLSX_H

#include "ImportSpreadsheet.h"

#include "eible_global.h"

class QuaZip;
class QuaZipFile;
class QXmlStreamReader;

class EIBLE_EXPORT ImportXlsx : public ImportSpreadsheet
{
public:
    explicit ImportXlsx(QIODevice& ioDevice);

    std::pair<bool, QMap<QString, QString>> getSheetList() override;

    std::pair<bool, QStringList> getColumnList(
        const QString& sheetPath,
        const QHash<QString, int>& sharedStrings) override;

    std::tuple<bool, QList<int>, QList<int>> getStyles();

    std::pair<bool, QStringList> getSharedStrings();

    std::pair<bool, QVector<ColumnType>> getColumnTypes(
        const QString& sheetPath, int columnsCount,
        const QHash<QString, int>& sharedStrings, const QList<int>& dateStyles,
        const QList<int>& allStyles) override;

private:
    bool openZipAndMoveToSecondRow(QuaZip& zip, const QString& sheetName,
                                   QuaZipFile& zipFile,
                                   QXmlStreamReader& xmlStreamReader);

    static constexpr int NOT_SET_COLUMN{-1};

    static constexpr int DECIMAL_BASE{10};
};

#endif  // IMPORTXLSX_H
