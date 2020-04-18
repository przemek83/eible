#ifndef IMPORTXLSX_H
#define IMPORTXLSX_H

#include "ImportSpreadsheet.h"

#include "eible_global.h"

class EIBLE_EXPORT ImportXlsx : public ImportSpreadsheet
{
public:
    explicit ImportXlsx(QIODevice& ioDevice);

    std::pair<bool, QMap<QString, QString>> getSheetList() override;

    std::pair<bool, QStringList> getColumnList(
        const QString& sheetName, QHash<QString, int> sharedStrings) override;

    std::tuple<bool, QList<int>, QList<int>> getStyles();

    std::pair<bool, QStringList> getSharedStrings();
};

#endif  // IMPORTXLSX_H
