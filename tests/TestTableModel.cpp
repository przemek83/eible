#include "TestTableModel.h"

#include <QDate>

TestTableModel::TestTableModel(int columnCount, int rowCount)
    : QAbstractTableModel(), columnCount_(columnCount), rowCount_(rowCount)
{
    data_.resize(columnCount);
    for (int column = 0; column < columnCount; ++column)
    {
        data_[column].resize(rowCount);
        for (int row = 0; row < rowCount; ++row)
        {
            switch (column % 3)
            {
                case 0:
                    data_[column][row] =
                        QString("Item %1, %2").arg(column).arg(row);
                    break;
                case 1:
                    data_[column][row] = column + row;
                    break;
                case 2:
                    const QDate date(2020, 1, 1);
                    data_[column][row] = date.addDays(column + row);
                    break;
            }
        }
    }
}

int TestTableModel::rowCount([[maybe_unused]] const QModelIndex& parent) const
{
    return rowCount_;
}

int TestTableModel::columnCount([[maybe_unused]] const QModelIndex& parent) const
{
    return columnCount_;
}

QVariant TestTableModel::headerData(int section,
                                    [[maybe_unused]] Qt::Orientation orientation,
                                    [[maybe_unused]] int role) const
{
    const QStringList headers {"Text", "Numeric", "Date"};
    return headers[section % 3];
}
