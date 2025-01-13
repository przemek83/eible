#include <eible/ImportSpreadsheet.h>

#include <quazip/quazipfile.h>
#include <QCoreApplication>
#include <QIODevice>
#include <QVariant>
#include <QVector>

#include "Utilities.h"

ImportSpreadsheet::ImportSpreadsheet(QIODevice& ioDevice)
    : ioDevice_{ioDevice}, zip_{&ioDevice_}
{
}

QString ImportSpreadsheet::getLastError() const { return lastError_; }

void ImportSpreadsheet::setNameForEmptyColumn(const QString& name)
{
    emptyColName_ = name;
}

std::pair<bool, int> ImportSpreadsheet::getColumnCount(const QString& sheetName)
{
    return getCount(sheetName, columnCounts_);
}

std::pair<bool, int> ImportSpreadsheet::getRowCount(const QString& sheetName)
{
    return getCount(sheetName, rowCounts_);
}

std::pair<bool, QVector<QVector<QVariant> > > ImportSpreadsheet::getData(
    const QString& sheetName, const QVector<int>& excludedColumns)
{
    auto [success, rowCount]{getRowCount(sheetName)};
    if (!success)
        return {false, {}};
    return getLimitedData(sheetName, excludedColumns, rowCount);
}

void ImportSpreadsheet::setError(const QString& errorContent)
{
    lastError_ = errorContent;
}

void ImportSpreadsheet::updateProgress(int currentRow, int rowCount,
                                       int& lastEmittedPercent)
{
    const int currentPercent{
        static_cast<int>((100. * (currentRow + 1)) / rowCount)};
    if (currentPercent > lastEmittedPercent)
    {
        Q_EMIT progressPercentChanged(currentPercent);
        lastEmittedPercent = currentPercent;
        QCoreApplication::processEvents();
    }
}

bool ImportSpreadsheet::openZip()
{
    if ((!zip_.isOpen()) && (!zip_.open(QuaZip::mdUnzip)))
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

bool ImportSpreadsheet::openZipFile(QuaZipFile& quaZipFile)
{
    quaZipFile.setZip(&zip_);
    if (!quaZipFile.open(QIODevice::ReadOnly | QIODevice::Text))
    {
        setError("Can not open file " + quaZipFile.getFileName() + ".");
        return false;
    }
    return true;
}

bool ImportSpreadsheet::initZipFile(QuaZipFile& quaZipFile,
                                    const QString& zipFileName)
{
    return openZip() && setCurrentZipFile(zipFileName) &&
           openZipFile(quaZipFile);
}

QMap<int, int> ImportSpreadsheet::createActiveColumnMapping(
    const QVector<int>& excludedColumns, int columnCount)
{
    QMap<int, int> activeColumnsMapping;
    int columnToFill{0};
    for (int i{0}; i < columnCount; ++i)
    {
        if (!excludedColumns.contains(i))
        {
            activeColumnsMapping[i] = columnToFill;
            ++columnToFill;
        }
    }
    return activeColumnsMapping;
}

QVector<QVariant> ImportSpreadsheet::createTemplateDataRow(
    const QVector<int>& excludedColumns, const QVector<ColumnType>& columnTypes)
{
    QVector<QVariant> templateDataRow;
    int columnToFill{0};
    templateDataRow.resize(columnTypes.size() - excludedColumns.size());

    const qsizetype columnCount{columnTypes.size()};
    for (qsizetype i{0}; i < columnCount; ++i)
    {
        if (!excludedColumns.contains(i))
        {
            templateDataRow[columnToFill] =
                utilities::getDefaultVariantForFormat(columnTypes[i]);
            ++columnToFill;
        }
    }
    return templateDataRow;
}

bool ImportSpreadsheet::columnsToExcludeAreValid(
    const QVector<int>& excludedColumns, int columnCount)
{
    if (const auto it{std::find_if(
            excludedColumns.begin(), excludedColumns.end(),
            [count = columnCount](int column) { return column >= count; })};
        it != excludedColumns.end())
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
    auto [namesRetrieved, columnNames]{retrieveColumnNames(sheetName)};
    if (!namesRetrieved)
        return false;

    auto [success, rowCount,
          columnTypes]{retrieveRowCountAndColumnTypes(sheetName)};
    if (!success)
        return false;

    const qsizetype actualColumnCount{
        std::max(columnNames.size(), columnTypes.size())};
    columnNames.reserve(actualColumnCount);
    columnTypes.reserve(actualColumnCount);

    for (qsizetype i{columnNames.size()}; i < actualColumnCount; ++i)
        columnNames << emptyColName_;

    for (qsizetype i{columnTypes.size()}; i < actualColumnCount; ++i)
        columnTypes.append(ColumnType::STRING);

    rowCounts_[sheetName] = rowCount;
    columnCounts_[sheetName] = static_cast<int>(actualColumnCount);
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
