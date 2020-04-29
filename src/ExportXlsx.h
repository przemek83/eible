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
    bool writeContent(const QByteArray& content, QIODevice& ioDevice) override;

    QByteArray getEmptyContent() override;

    QByteArray generateHeaderContent(const QAbstractItemModel& model) override;

    QByteArray generateRowContent(const QAbstractItemModel& model, int row,
                                  int skippedRowsCount) override;

    QByteArray getContentEnding() override;

private:
    const QByteArray& getCellTypeTag(QVariant& cell);

    QList<QByteArray> columnNames_{};
};

#endif  // EXPORTXLSX_H
