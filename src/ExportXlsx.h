#ifndef EXPORTXLSX_H
#define EXPORTXLSX_H

#include <QVector>

#include "ExportData.h"
#include "eible_global.h"

class QAbstractItemModel;
class QAbstractItemView;
class QIODevice;

/**
 * @class ExportXlsx
 * @brief Class for exporting data to xlsx files.
 */
class EIBLE_EXPORT ExportXlsx : public ExportData
{
    Q_OBJECT
public:
    explicit ExportXlsx(QObject* parent = nullptr);
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
    const QByteArray& getCellTypeTag(const QVariant& cell);

    void initColumnNames(int modelColumnCount);

    const QByteArray CELL_START{QByteArrayLiteral("<c r=\"")};
    const QByteArray ROW_NUMBER_CLOSE{QByteArrayLiteral("\" ")};
    const QByteArray VALUE_START{QByteArrayLiteral("><v>")};
    const QByteArray CELL_END{QByteArrayLiteral("</v></c>")};
    const QByteArray DATE_TAG{QByteArrayLiteral("s=\"3\"")};
    const QByteArray NUMERIC_TAG{QByteArrayLiteral("s=\"4\"")};
    const QByteArray STRING_TAG{QByteArrayLiteral("t=\"str\"")};

    QVector<QByteArray> columnNames_{};
};

#endif  // EXPORTXLSX_H
