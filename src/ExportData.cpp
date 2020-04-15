#include "ExportData.h"

bool ExportData::exportView(const QAbstractItemView& view, QIODevice& ioDevice)
{
    QByteArray content{generateContent(view)};
    return writeContent(content, ioDevice);
}
