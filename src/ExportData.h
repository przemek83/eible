#ifndef EXPORT_H
#define EXPORT_H

#include <QObject>

#include "eible_global.h"

class QAbstractItemView;
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

    virtual bool exportView(const QAbstractItemView& view,
                            QIODevice& ioDevice) = 0;
};

#endif  // EXPORT_H
