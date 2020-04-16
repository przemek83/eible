#include "ImportSpreadsheet.h"

#include <QIODevice>

ImportSpreadsheet::ImportSpreadsheet(QIODevice& ioDevice) : ioDevice_(ioDevice)
{
}

std::pair<QString, QString> ImportSpreadsheet::getError() const
{
    return error_;
}

void ImportSpreadsheet::setError(QString functionName, QString errorContent)
{
    error_ = {functionName, errorContent};
}
