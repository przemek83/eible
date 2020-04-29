#ifndef EXPORTDATA_H
#define EXPORTDATA_H

#include <QObject>

#include "eible_global.h"

class QAbstractItemView;
class QAbstractItemModel;
class QIODevice;

class EIBLE_EXPORT ExportData : public QObject
{
    Q_OBJECT
public:
    ExportData() = default;
    ~ExportData() override = default;

    ExportData& operator=(const ExportData& other) = delete;
    ExportData(const ExportData& other) = delete;

    ExportData& operator=(ExportData&& other) = delete;
    ExportData(ExportData&& other) = delete;

    bool exportView(const QAbstractItemView& view, QIODevice& ioDevice);

protected:
    virtual bool writeContent(const QByteArray& content,
                              QIODevice& ioDevice) = 0;

    bool rowShouldBeSkipped(const QAbstractItemView& view, int row);

    virtual QByteArray getEmptyContent() = 0;

    virtual QByteArray generateHeaderContent(
        const QAbstractItemModel& model) = 0;

    virtual QByteArray generateRowContent(const QAbstractItemModel& model,
                                          int row, int skippedRowsCount) = 0;

    virtual QByteArray getContentEnding();

    virtual void updateProgress(unsigned int currentRow, unsigned int rowCount,
                                unsigned int& lastEmittedPercent);

Q_SIGNALS:
    void progressPercentChanged(unsigned int currentPercent);
};

#endif  // EXPORTDATA_H
