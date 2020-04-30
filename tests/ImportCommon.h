#ifndef IMPORTCOMMON_H
#define IMPORTCOMMON_H

#include <ImportSpreadsheet.h>
#include <QStringList>

class ImportCommon
{
public:
    template <class T>
    void checkRetrievingSheetNames(const QString& fileName);

    template <class T>
    void checkRetrievingSheetNamesFromEmptyFile();

private:
    static const QStringList sheetNames_;
};

#endif  // IMPORTCOMMON_H
