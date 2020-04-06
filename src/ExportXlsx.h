#ifndef EXPORTXLSX_H
#define EXPORTXLSX_H

#include <QObject>

#include "eible_global.h"

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

    bool exportView(const QAbstractItemView* view, QIODevice* ioDevice);

private:
    const QString& getCellTypeTag(QVariant& cell);

    QByteArray gatherSheetContent(const QAbstractItemView* view);

Q_SIGNALS:
    void updateProgress(int progress);
};

#endif // EXPORTXLSX_H
