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
    explicit ExportXlsx(const QString& filePath);

    bool exportView(const QAbstractItemView* view, QIODevice* ioDevice);

private:
    const QString& getCellTypeTag(QVariant& cell);

    QByteArray gatherSheetContent(const QAbstractItemView* view);

    const QString filePath_;

Q_SIGNALS:
    void updateProgress(int progress);
};

#endif // EXPORTXLSX_H
