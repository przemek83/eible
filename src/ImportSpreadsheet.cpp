#include "ImportSpreadsheet.h"

#include <QCoreApplication>
#include <QIODevice>

ImportSpreadsheet::ImportSpreadsheet(QIODevice& ioDevice)
    : ioDevice_(ioDevice), emptyColName_("---")
{
}

std::pair<QString, QString> ImportSpreadsheet::getError() const
{
    return error_;
}

void ImportSpreadsheet::setNameForEmptyColumn(const QString& name)
{
    emptyColName_ = name;
}

void ImportSpreadsheet::setError(QString functionName, QString errorContent)
{
    error_ = {functionName, errorContent};
}

void ImportSpreadsheet::updateProgress(unsigned int rowCounter,
                                       unsigned int rowLimit,
                                       unsigned int& lastEmittedPercent)
{
    const unsigned int currentPercent{
        static_cast<unsigned int>(100. * rowCounter / rowLimit)};
    if (currentPercent > lastEmittedPercent)
    {
        emit progressPercentChanged(currentPercent);
        lastEmittedPercent = currentPercent;
        QCoreApplication::processEvents();
    }
}
