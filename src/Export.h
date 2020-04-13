#ifndef EXPORT_H
#define EXPORT_H

#include <QObject>

class QAbstractItemView;
class QIODevice;

class Export : public QObject
{
    Q_OBJECT
public:
    Export() = default;
    ~Export() override = default;

    Export& operator=(const Export& other) = delete;
    Export(const Export& other) = delete;

    Export& operator=(Export&& other) = delete;
    Export(Export&& other) = delete;

    virtual bool exportView(const QAbstractItemView& view,
                            QIODevice& ioDevice) = 0;
};

#endif  // EXPORT_H
