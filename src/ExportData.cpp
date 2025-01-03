#include <eible/ExportData.h>

#include <QAbstractItemView>
#include <QCoreApplication>

ExportData::ExportData(QObject* parent) : QObject(parent) {}

bool ExportData::exportView(const QAbstractItemView& view, QIODevice& ioDevice)
{
    const QAbstractItemModel* model{view.model()};
    Q_ASSERT(model != nullptr);

    if (model->columnCount() == 0)
        return writeContent(getEmptyContent(), ioDevice);

    QByteArray content{generateHeaderContent(*model)};
    int skippedRows{0};
    int lastEmittedPercent{0};
    const int rowCount{model->rowCount()};
    for (int row{0}; row < rowCount; ++row)
    {
        if (rowShouldBeSkipped(view, row))
        {
            ++skippedRows;
        }
        else
        {
            content.append(generateRowContent(*model, row, skippedRows));
            updateProgress(row, rowCount, lastEmittedPercent);
        }
    }
    content.append(getContentEnding());
    return writeContent(content, ioDevice);
}

bool ExportData::rowShouldBeSkipped(const QAbstractItemView& view, int row)
{
    const QAbstractItemModel* model{view.model()};
    const bool multiSelection{
        (QAbstractItemView::MultiSelection == view.selectionMode())};
    const QItemSelectionModel* selectionModel{view.selectionModel()};

    return multiSelection &&
           (!selectionModel->isSelected(model->index(row, 0)));
}

QByteArray ExportData::getContentEnding() { return {}; }

void ExportData::updateProgress(int currentRow, int rowCount,
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
