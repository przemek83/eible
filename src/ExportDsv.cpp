#include <eible/ExportDsv.h>

#include <QAbstractItemView>
#include <QDate>
#include <QIODevice>
#include <QVariant>

ExportDsv::ExportDsv(char separator, QObject* parent)
    : ExportData(parent), separator_{separator}
{
}

ExportDsv::ExportDsv(char separator) : ExportDsv(separator, nullptr) {}

QByteArray ExportDsv::getEmptyContent() { return {}; }

QByteArray ExportDsv::generateHeaderContent(const QAbstractItemModel& model)
{
    QByteArray headersContent;
    const int columnCount{model.columnCount()};
    for (int j{0}; j < columnCount; ++j)
    {
        headersContent.append(
            model.headerData(j, Qt::Horizontal).toString().toUtf8());
        if (j != (columnCount - 1))
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
    const int columnCount{model.columnCount()};
    for (int j{0}; j < columnCount; ++j)
    {
        if (const QVariant actualField{model.index(row, j).data()};
            !actualField.isNull())
            variantToString(actualField, rowContent, separator_);
        if (j != (columnCount - 1))
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

QByteArray ExportDsv::convertToByteArray(const QVariant& variant,
                                         char separator)
{
    // Following https://tools.ietf.org/html/rfc4180
    QByteArray value{variant.toByteArray()};
    if (value.contains(separator) || value.contains('\"') ||
        value.contains('\n'))
    {
        value.replace('"', QByteArrayLiteral("\"\""));
        return "\"" + value + "\"";
    }

    return value;
}

void ExportDsv::variantToString(const QVariant& variant, QByteArray& array,
                                char separator) const
{
    switch (variant.typeId())
    {
        case QMetaType::Double:
        case QMetaType::Int:
        {
            array.append(
                locale_.toString(variant.toDouble(), 'f', precision_).toUtf8());
            break;
        }

        case QMetaType::QDate:
        case QMetaType::QDateTime:
        {
            if (dateFormat_.isEmpty())
                array.append(variant.toDate().toString(qtDateFormat_).toUtf8());
            else
                array.append(variant.toDate().toString(dateFormat_).toUtf8());
            break;
        }

        case QMetaType::QString:
        {
            const QByteArray byteArray{convertToByteArray(variant, separator)};
            array.append(byteArray);
            break;
        }

        default:
        {
            Q_ASSERT(false);
            break;
        }
    }
}
