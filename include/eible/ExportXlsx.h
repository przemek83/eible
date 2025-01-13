#pragma once

#include <QVector>

#include "ExportData.h"
#include "eible_global.h"

class QAbstractItemModel;
class QAbstractItemView;
class QIODevice;

/// @class ExportXlsx
/// @brief Class for exporting data to xlsx files.
class EIBLE_EXPORT ExportXlsx : public ExportData
{
    Q_OBJECT
public:
    explicit ExportXlsx(QObject* parent = nullptr);

protected:
    bool writeContent(const QByteArray& content, QIODevice& ioDevice) override;

    QByteArray getEmptyContent() override;

    QByteArray generateHeaderContent(const QAbstractItemModel& model) override;

    QByteArray generateRowContent(const QAbstractItemModel& model, int row,
                                  int skippedRowsCount) override;

    QByteArray getContentEnding() override;

private:
    const QByteArray& getCellTypeTag(const QVariant& cell) const;

    void initColumnNames(int modelColumnCount);

    const QByteArray cellStart_{QByteArrayLiteral("<c r=\"")};
    const QByteArray rowNumberClose_{QByteArrayLiteral("\" ")};
    const QByteArray valueStart_{QByteArrayLiteral("><v>")};
    const QByteArray cellEnd_{QByteArrayLiteral("</v></c>")};
    const QByteArray dateTag_{QByteArrayLiteral("s=\"3\"")};
    const QByteArray numericTag_{QByteArrayLiteral("s=\"4\"")};
    const QByteArray stringTag_{QByteArrayLiteral("t=\"str\"")};

    QVector<QByteArray> columnNames_{};
};
