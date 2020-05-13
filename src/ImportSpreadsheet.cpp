#include "ImportSpreadsheet.h"

#include <Qt5Quazip/quazipfile.h>
#include <QCoreApplication>
#include <QIODevice>
#include <QVariant>
#include <QVector>

#include "EibleUtilities.h"

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

bool ImportSpreadsheet::openZip()
{
    if (!zip_.isOpen() && !zip_.open(QuaZip::mdUnzip))
    {
        setError(__FUNCTION__,
                 "Can not open zip file " + zip_.getZipName() + ".");
        return false;
    }
    return true;
}

bool ImportSpreadsheet::setCurrentZipFile(const QString& zipFileName)
{
    if (!zip_.setCurrentFile(zipFileName))
    {
        setError(__FUNCTION__,
                 "Can not find file " + zipFileName + " in archive.");
        return false;
    }
    return true;
}

bool ImportSpreadsheet::openZipFile(QuaZipFile& zipFile)
{
    zipFile.setZip(&zip_);
    if (!zipFile.open(QIODevice::ReadOnly | QIODevice::Text))
    {
        setError(__FUNCTION__,
                 "Can not open file " + zipFile.getFileName() + ".");
        return false;
    }
    return true;
}

bool ImportSpreadsheet::initZipFile(QuaZipFile& zipFile,
                                    const QString& zipFileName)
{
    return openZip() && setCurrentZipFile(zipFileName) && openZipFile(zipFile);
}

QMap<unsigned int, unsigned int> ImportSpreadsheet::createActiveColumnMapping(
    const QVector<unsigned int>& excludedColumns,
    unsigned int columnCount) const
{
    QMap<unsigned int, unsigned int> activeColumnsMapping;
    unsigned int columnToFill{0};
    for (unsigned int i = 0; i < columnCount; ++i)
    {
        if (!excludedColumns.contains(i))
        {
            activeColumnsMapping[i] = columnToFill;
            columnToFill++;
        }
    }
    return activeColumnsMapping;
}

QVector<QVariant> ImportSpreadsheet::createTemplateDataRow(
    const QVector<unsigned int>& excludedColumns,
    const QVector<ColumnType>& columnTypes) const
{
    QVector<QVariant> templateDataRow;
    int columnToFill = 0;
    templateDataRow.resize(columnTypes.size() - excludedColumns.size());
    for (unsigned int i = 0; i < static_cast<unsigned int>(columnTypes.size());
         ++i)
    {
        if (!excludedColumns.contains(i))
        {
            templateDataRow[columnToFill] =
                EibleUtilities::getDefaultVariantForFormat(
                    columnTypes[static_cast<int>(i)]);
            columnToFill++;
        }
    }
    return templateDataRow;
}

bool ImportSpreadsheet::columnsToExcludeAreValid(
    const QVector<unsigned int>& excludedColumns, unsigned int columnCount)
{
    auto it = std::find_if(
        excludedColumns.begin(), excludedColumns.end(),
        [=](unsigned int column) { return column >= columnCount; });
    if (it != excludedColumns.end())
    {
        setError(__FUNCTION__, "Column to exclude " + QString::number(*it) +
                                   " is invalid. Xlsx got only " +
                                   QString::number(columnCount) +
                                   " columns indexed from 0.");
        return false;
    }
    return true;
}

bool ImportSpreadsheet::analyzeSheet(const QString& sheetName)
{
    auto [namesRetrieved, columnNames] = retrieveColumnNames(sheetName);
    if (!namesRetrieved)
        return false;

    auto [success, rowCount, columnTypes] =
        retrieveRowCountAndColumnTypes(sheetName);
    if (!success)
        return false;

    const unsigned int actualColumnCount{static_cast<unsigned int>(
        std::max(columnNames.size(), columnTypes.size()))};

    for (unsigned int i = static_cast<unsigned int>(columnNames.size());
         i < actualColumnCount; ++i)
        columnNames << emptyColName_;

    for (unsigned int i = static_cast<unsigned int>(columnTypes.size());
         i < actualColumnCount; ++i)
        columnTypes.append(ColumnType::STRING);

    rowCounts_[sheetName] = rowCount;
    columnCounts_[sheetName] = actualColumnCount;
    columnTypes_[sheetName] = columnTypes;
    columnNames_[sheetName] = columnNames;

    return true;
}

bool ImportSpreadsheet::isSheetNameValid(const QStringList& sheetNames,
                                         const QString& sheetName)
{
    if (!sheetNames.contains(sheetName))
    {
        setError(__FUNCTION__,
                 "Sheet " + sheetName +
                     " not found. Available sheets: " + sheetNames.join(','));
        return false;
    }
    return true;
}
