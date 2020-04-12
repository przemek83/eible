#ifndef EXPORTXLSX_H
#define EXPORTXLSX_H

#include <QObject>

#include "eible_global.h"

class QAbstractItemModel;
class QAbstractItemView;
class QIODevice;

class EIBLE_EXPORT ExportXlsx : public QObject
{
    Q_OBJECT
public:
    ExportXlsx() = default;
    ~ExportXlsx() override = default;

    ExportXlsx& operator=(const ExportXlsx& other) = delete;
    ExportXlsx(const ExportXlsx& other) = delete;

    ExportXlsx& operator=(ExportXlsx&& other) = delete;
    ExportXlsx(ExportXlsx&& other) = delete;

    bool exportView(const QAbstractItemView& view, QIODevice& ioDevice);

private:
    const QByteArray& getCellTypeTag(QVariant& cell);

    QByteArray gatherSheetContent(const QAbstractItemView& view);

    void addHeaders(QByteArray& rowsContent,
                    const QAbstractItemModel& proxyModel,
                    const QStringList& columnNames) const;

Q_SIGNALS:
    void updateProgress(int progress);
};

#endif  // EXPORTXLSX_H
