#ifndef TESTTABLEMODEL_H
#define TESTTABLEMODEL_H

#include <QAbstractTableModel>

class TestTableModel : public QAbstractTableModel
{
    Q_OBJECT
public:
    TestTableModel(int columnCount, int rowCount);
    ~TestTableModel() override = default;

    int rowCount([[maybe_unused]] const QModelIndex& parent =
                     QModelIndex()) const override;

    int columnCount([[maybe_unused]] const QModelIndex& parent =
                        QModelIndex()) const override;

    inline QVariant data(
        const QModelIndex& index,
        [[maybe_unused]] int role = Qt::DisplayRole) const override
    {
        return data_[index.column()][index.row() % MAX_DATA_ROWS];
    }

    QVariant headerData(
        int section, [[maybe_unused]] Qt::Orientation orientation,
        [[maybe_unused]] int role = Qt::DisplayRole) const override;

private:
    const int columnCount_;
    const int rowCount_;

    static constexpr int MAX_DATA_ROWS{100};

    QVector<QVector<QVariant>> data_;
};

#endif  // TESTTABLEMODEL_H
