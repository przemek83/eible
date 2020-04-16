#ifndef IMPORTXLSX_H
#define IMPORTXLSX_H

#include "ImportSpreadsheet.h"

class ImportXlsx : public ImportSpreadsheet
{
public:
    explicit ImportXlsx(QIODevice& ioDevice);

    std::pair<bool, QMap<QString, QString>> getSheetList() override;
};

#endif  // IMPORTXLSX_H
