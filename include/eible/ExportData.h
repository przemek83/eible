#pragma once

#include <QObject>

#include "eible_global.h"

class QAbstractItemView;
class QAbstractItemModel;
class QIODevice;

/// @class ExportData
/// @brief Base class for exporting classes.
class EIBLE_EXPORT ExportData : public QObject
{
    Q_OBJECT
public:
    explicit ExportData(QObject* parent);
    ExportData();

    /// @brief Export data from view to ioDevice.
    /// @param view Source of data.
    /// @param ioDevice Destination of exported content.
    /// @return True if success, false if failure.
    bool exportView(const QAbstractItemView& view, QIODevice& ioDevice);

protected:
    virtual bool writeContent(const QByteArray& content,
                              QIODevice& ioDevice) = 0;

    virtual QByteArray getEmptyContent() = 0;

    virtual QByteArray generateHeaderContent(
        const QAbstractItemModel& model) = 0;

    virtual QByteArray generateRowContent(const QAbstractItemModel& model,
                                          int row, int skippedRowsCount) = 0;

    virtual QByteArray getContentEnding();

    virtual void updateProgress(int currentRow, int rowCount,
                                int& lastEmittedPercent);

private:
    static bool rowShouldBeSkipped(const QAbstractItemView& view, int row);

Q_SIGNALS:
    /// Triggered on change of progress percentage.
    /// @param currentPercent New progress percent.
    void progressPercentChanged(int currentPercent);
};
