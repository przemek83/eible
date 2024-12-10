#include <eible/ImportSpreadsheet.h>

#include <quazip/quazipfile.h>
#include <QCoreApplication>
#include <QIODevice>
#include <QVariant>
#include <QVector>

#include "EibleUtilities.h"

ImportSpreadsheet::ImportSpreadsheet(QIODevice& ioDevice)
    : ioDevice_(ioDevice),
      emptyColName_(QStringLiteral("---")),
      zip_(&ioDevice_)
{
}

QString ImportSpreadsheet::getLastError() const { return lastError_; }

void ImportSpreadsheet::setNameForEmptyColumn(const QString& name)
{
    emptyColName_ = name;
}

std::pair<bool, unsigned int> ImportSpreadsheet::getColumnCount(
    const QString& sheetName)
{
    return getCount(sheetName, columnCounts_);
}

std::pair<bool, unsigned int> ImportSpreadsheet::getRowCount(
    const QString& sheetName)
{
    return getCount(sheetName, rowCounts_);
}

std::pair<bool, QVector<QVector<QVariant> > > ImportSpreadsheet::getData(
    const QString& sheetName, const QVector<unsigned int>& excludedColumns)
{
    auto [success, rowCount] = getRowCount(sheetName);
    if (!success)
        return {false, {}};
    return getLimitedData(sheetName, excludedColumns, rowCount);
}

void ImportSpreadsheet::setError(const QString& errorContent)
{
    lastError_ = errorContent;
}

void ImportSpreadsheet::updateProgress(unsigned int currentRow,
                                       unsigned int rowCount,
                                       unsigned int& lastEmittedPercent)
{
    const unsigned int currentPercent{
        static_cast<unsigned int>(100. * (currentRow + 1) / rowCount)};
    if (currentPercent > lastEmittedPercent)
    {
        Q_EMIT progressPercentChanged(currentPercent);
        lastEmittedPercent = currentPercent;
        QCoreApplication::processEvents();
    }
}

bool ImportSpreadsheet::openZip()
{
    if (!zip_.isOpen() && !zip_.open(QuaZip::mdUnzip))
    {
        setError("Can not open zip file " + zip_.getZipName() + ".");
        return false;
    }
    return true;
}

bool ImportSpreadsheet::setCurrentZipFile(const QString& zipFileName)
{
    if (!zip_.setCurrentFile(zipFileName))
    {
        setError("Can not find file " + zipFileName + " in archive.");
        return false;
    }
    return true;
}

bool ImportSpreadsheet::openZipFile(QuaZipFile& zipFile)
{
    zipFile.setZip(&zip_);
    if (!zipFile.open(QIODevice::ReadOnly | QIODevice::Text))
    {
        setError("Can not open file " + zipFile.getFileName() + ".");
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
    const QVector<unsigned int>& excludedColumns, unsigned int columnCount)
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
    const QVector<ColumnType>& columnTypes)
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
    const auto it = std::find_if(excludedColumns.begin(), excludedColumns.end(),
                                 [=](unsigned int column)
                                 { return column >= columnCount; });
    if (it != excludedColumns.end())
    {
        setError("Column to exclude " + QString::number(*it) +
                 " is invalid. Xlsx got only " + QString::number(columnCount) +
                 " columns (indexed from 0).");
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
    columnNames.reserve(static_cast<int>(actualColumnCount));
    columnTypes.reserve(static_cast<int>(actualColumnCount));

    for (auto i = static_cast<unsigned int>(columnNames.size());
         i < actualColumnCount; ++i)
        columnNames << emptyColName_;

    for (auto i = static_cast<unsigned int>(columnTypes.size());
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
        setError("Sheet " + sheetName +
                 " not found. Available sheets: " + sheetNames.join(','));
        return false;
    }
    return true;
}
