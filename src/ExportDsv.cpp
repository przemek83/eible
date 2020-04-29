#include "ExportDsv.h"

#include <QAbstractItemView>
#include <QDate>
#include <QVariant>

ExportDsv::ExportDsv(char separator) : ExportData(), separator_(separator) {}

QByteArray ExportDsv::getEmptyContent() { return QByteArray(); }

QByteArray ExportDsv::generateHeaderContent(const QAbstractItemModel& model)
{
    QByteArray headersContent;
    for (int j = 0; j < model.columnCount(); ++j)
    {
        headersContent.append(model.headerData(j, Qt::Horizontal).toString());
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
        QVariant actualField = model.index(row, j).data();
        if (!actualField.isNull())
            variantToString(actualField, rowContent, separator_);
        if (j != model.columnCount() - 1)
            rowContent.append(separator_);
    }

    return rowContent;
}

void ExportDsv::setDateFormat(Qt::DateFormat format) { qtDateFormat_ = format; }

void ExportDsv::setDateFormat(QString format) { dateFormat_ = format; }

void ExportDsv::setNumbersLocale(QLocale locale) { locale_ = locale; }

bool ExportDsv::writeContent(const QByteArray& content, QIODevice& ioDevice)
{
    return ioDevice.write(content) != -1;
}

QString doubleToStringUsingLocale(double value, int precision)
{
    static bool initialized{false};
    static QLocale locale;
    if (!initialized)
    {
        locale.setNumberOptions(locale.numberOptions() &
                                ~QLocale::OmitGroupSeparator);
        initialized = true;
    }

    return locale.toString(value, 'f', precision);
}

void ExportDsv::variantToString(const QVariant& variant,
                                QByteArray& destinationArray,
                                char separator) const
{
    switch (variant.type())
    {
        case QVariant::Double:
        case QVariant::Int:
        {
            destinationArray.append(
                locale_.toString(variant.toDouble(), 'f', 2));
            break;
        }

        case QVariant::Date:
        case QVariant::DateTime:
        {
            if (dateFormat_.isEmpty())
                destinationArray.append(
                    variant.toDate().toString(qtDateFormat_));
            else
                destinationArray.append(variant.toDate().toString(dateFormat_));
            break;
        }

        case QVariant::String:
        {
            // Following https://tools.ietf.org/html/rfc4180
            QByteArray value{variant.toByteArray()};
            if (value.contains(separator) || value.contains('\"') ||
                value.contains('\n'))
            {
                value.replace('"', QStringLiteral("\"\""));
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
