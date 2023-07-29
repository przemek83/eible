#include "ExportDsv.h"

#include <QAbstractItemView>
#include <QDate>
#include <QIODevice>
#include <QVariant>

ExportDsv::ExportDsv(char separator, QObject* parent)
    : ExportData(parent), separator_(separator)
{
}

QByteArray ExportDsv::getEmptyContent() { return {}; }

QByteArray ExportDsv::generateHeaderContent(const QAbstractItemModel& model)
{
    QByteArray headersContent;
    for (int j = 0; j < model.columnCount(); ++j)
    {
        headersContent.append(
            model.headerData(j, Qt::Horizontal).toString().toUtf8());
        if (j != model.columnCount() - 1)
            headersContent.append(separator_);
    }
    return headersContent;
}

QByteArray ExportDsv::generateRowContent(const QAbstractItemModel& model,
                                         int row,
                                         [[maybe_unused]] int skippedRowsCount)
{
    QByteArray rowContent;
    rowContent.append("\n");
    for (int j = 0; j < model.columnCount(); ++j)
    {
        const QVariant actualField = model.index(row, j).data();
        if (!actualField.isNull())
            variantToString(actualField, rowContent, separator_);
        if (j != model.columnCount() - 1)
            rowContent.append(separator_);
    }

    return rowContent;
}

void ExportDsv::setDateFormat(Qt::DateFormat format) { qtDateFormat_ = format; }

void ExportDsv::setDateFormat(const QString& format) { dateFormat_ = format; }

void ExportDsv::setNumbersLocale(const QLocale& locale) { locale_ = locale; }

bool ExportDsv::writeContent(const QByteArray& content, QIODevice& ioDevice)
{
    return ioDevice.write(content) != -1;
}

void ExportDsv::variantToString(const QVariant& variant,
                                QByteArray& destinationArray,
                                char separator) const
{
    switch (variant.typeId())
    {
        case QMetaType::Double:
        case QMetaType::Int:
        {
            destinationArray.append(
                locale_.toString(variant.toDouble(), 'f', 2).toUtf8());
            break;
        }

        case QMetaType::QDate:
        case QMetaType::QDateTime:
        {
            if (dateFormat_.isEmpty())
                destinationArray.append(
                    variant.toDate().toString(qtDateFormat_).toUtf8());
            else
                destinationArray.append(
                    variant.toDate().toString(dateFormat_).toUtf8());
            break;
        }

        case QMetaType::QString:
        {
            // Following https://tools.ietf.org/html/rfc4180
            QByteArray value{variant.toByteArray()};
            if (value.contains(separator) || value.contains('\"') ||
                value.contains('\n'))
            {
                value.replace('"', QByteArrayLiteral("\"\""));
                destinationArray.append("\"" + value + "\"");
            }
            else
                destinationArray.append(value);
            break;
        }

        default:
        {
            Q_ASSERT(false);
            break;
        }
    }
}
