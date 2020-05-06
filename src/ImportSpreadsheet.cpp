#include "ImportSpreadsheet.h"

#include <QCoreApplication>
#include <QIODevice>
#include <QVariant>
#include <QVector>

#include <quazip5/quazipfile.h>

ImportSpreadsheet::ImportSpreadsheet(QIODevice& ioDevice)
    : ioDevice_(ioDevice), emptyColName_("---"), zip_(&ioDevice_)
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

bool ImportSpreadsheet::openZipFile(QuaZipFile& zipFile,
                                    const QString& zipFileName)
{
    if (!zip_.isOpen() && !zip_.open(QuaZip::mdUnzip))
    {
        setError(__FUNCTION__,
                 "Can not open zip file " + zip_.getZipName() + ".");
        return false;
    }

    if (!zip_.setCurrentFile(zipFileName))
    {
        setError(__FUNCTION__,
                 "Can not find file " + zipFileName + " in archive.");
        return false;
    }

    zipFile.setZip(&zip_);
    if (!zipFile.open(QIODevice::ReadOnly | QIODevice::Text))
    {
        setError(__FUNCTION__,
                 "Can not open file " + zipFile.getFileName() + ".");
        return false;
    }

    return true;
}
