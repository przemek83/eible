#include "ExportData.h"

#include <QAbstractItemView>
#include <QCoreApplication>

bool ExportData::exportView(const QAbstractItemView& view, QIODevice& ioDevice)
{
    const auto model = view.model();
    Q_ASSERT(model != nullptr);

    const int proxyColumnCount = model->columnCount();
    if (proxyColumnCount == 0)
        return writeContent(getEmptyContent(), ioDevice);

    QByteArray content{generateHeaderContent(*model)};
    int skippedRows = 0;
    unsigned int lastEmittedPercent{0};
    for (int row = 0; row < model->rowCount(); ++row)
    {
        if (rowShouldBeSkipped(view, row))
        {
            skippedRows++;
            continue;
        }
        content.append(generateRowContent(*model, row, skippedRows));
        updateProgress(row, model->rowCount(), lastEmittedPercent);
    }
    content.append(getContentEnding());
    return writeContent(content, ioDevice);
}

bool ExportData::rowShouldBeSkipped(const QAbstractItemView& view, int row)
{
    const auto model = view.model();
    const bool multiSelection =
        (QAbstractItemView::MultiSelection == view.selectionMode());
    QItemSelectionModel* selectionModel = view.selectionModel();

    return multiSelection && !selectionModel->isSelected(model->index(row, 0));
}

QByteArray ExportData::getContentEnding() { return QByteArray(); }

void ExportData::updateProgress(unsigned int currentRow, unsigned int rowCount,
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
