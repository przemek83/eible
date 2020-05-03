#include "ImportSpreadsheet.h"

#include <QCoreApplication>
#include <QIODevice>
#include <QVariant>
#include <QVector>

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

std::pair<bool, QVector<QVector<QVariant> > > ImportSpreadsheet::getData(
    const QString& sheetName, const QVector<unsigned int>& excludedColumns)
{
    auto [success, rowCount] = getRowCount(sheetName);
    if (!success)
        return {false, {}};
    return getLimitedData(sheetName, excludedColumns, rowCount);
}

void ImportSpreadsheet::setError(QString functionName, QString errorContent)
{
    error_ = {functionName, errorContent};
}

void ImportSpreadsheet::updateProgress(unsigned int currentRow,
                                       unsigned int rowCount,
                                       unsigned int& lastEmittedPercent)
{
    const unsigned int currentPercent{
        static_cast<unsigned int>(100. * (currentRow + 1) / rowCount)};
    if (currentPercent > lastEmittedPercent)
    {
        emit progressPercentChanged(currentPercent);
        lastEmittedPercent = currentPercent;
        QCoreApplication::processEvents();
    }
}
