#ifndef EXPORTXLSX_H
#define EXPORTXLSX_H

#include "ExportData.h"

#include "eible_global.h"

class QAbstractItemModel;
class QAbstractItemView;
class QIODevice;

class EIBLE_EXPORT ExportXlsx : public ExportData
{
    Q_OBJECT
public:
    ExportXlsx() = default;
    ~ExportXlsx() override = default;

    ExportXlsx& operator=(const ExportXlsx& other) = delete;
    ExportXlsx(const ExportXlsx& other) = delete;

    ExportXlsx& operator=(ExportXlsx&& other) = delete;
    ExportXlsx(ExportXlsx&& other) = delete;

protected:
    QByteArray generateContent(const QAbstractItemView& view) override;

    bool writeContent(const QByteArray& content, QIODevice& ioDevice) override;

private:
    const QByteArray& getCellTypeTag(QVariant& cell);

    void addHeaders(QByteArray& rowsContent,
                    const QAbstractItemModel& proxyModel,
                    const QStringList& columnNames) const;

Q_SIGNALS:
    void updateProgress(int progress);
};

#endif  // EXPORTXLSX_H
