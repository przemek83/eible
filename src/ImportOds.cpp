#include "ImportOds.h"

#include <QVector>
#include <utility>

ImportOds::ImportOds(QIODevice& ioDevice) : ImportSpreadsheet(ioDevice) {}

std::pair<bool, QStringList> ImportOds::getSheetNames() {}

std::pair<bool, QVector<ColumnType> > ImportOds::getColumnTypes(
    const QString& sheetName)
{
}

std::pair<bool, QStringList> ImportOds::getColumnNames(const QString& sheetName)
{
}

std::pair<bool, unsigned int> ImportOds::getColumnCount(
    const QString& sheetName)
{
}

std::pair<bool, unsigned int> ImportOds::getRowCount(const QString& sheetName)
{
}

std::pair<bool, QVector<QVector<QVariant> > > ImportOds::getData(
    const QString& sheetName, const QVector<unsigned int>& excludedColumns)
{
}

std::pair<bool, QVector<QVector<QVariant> > > ImportOds::getLimitedData(
    const QString& sheetName, const QVector<unsigned int>& excludedColumns,
    unsigned int rowLimit)
{
}
