#ifndef IMPORTXLSX_H
#define IMPORTXLSX_H

#include "ImportSpreadsheet.h"

#include "eible_global.h"

class EIBLE_EXPORT ImportXlsx : public ImportSpreadsheet
{
public:
    explicit ImportXlsx(QIODevice& ioDevice);

    std::pair<bool, QMap<QString, QString>> getSheetList() override;
};

#endif  // IMPORTXLSX_H
