#include "ExportDsv.h"

#include <QAbstractItemView>
#include <QDate>
#include <QVariant>

ExportDsv::ExportDsv(char separator) : ExportData(), separator_(separator) {}

QByteArray ExportDsv::generateContent(const QAbstractItemView& view)
{
    const auto proxyModel = view.model();
    Q_ASSERT(proxyModel != nullptr);

    int proxyColumnCount = proxyModel->columnCount();
    QByteArray destinationArray;
    if (proxyColumnCount == 0)
        return destinationArray;

    // Save column names.
    for (int j = 0; j < proxyColumnCount; ++j)
    {
        destinationArray.append(
            proxyModel->headerData(j, Qt::Horizontal).toString());
        if (j != proxyColumnCount - 1)
            destinationArray.append(separator_);
    }

    //    const QString barTitle =
    //        Constants::getProgressBarTitle(Constants::BarTitle::SAVING);
    //    ProgressBarCounter bar(barTitle, proxyModel->rowCount(), nullptr);
    //    bar.showDetached();

    bool multiSelection =
        (QAbstractItemView::MultiSelection == view.selectionMode());
    QItemSelectionModel* selectionModel = view.selectionModel();

    int proxyRowCount = proxyModel->rowCount();
    for (int i = 0; i < proxyRowCount; ++i)
    {
        if (multiSelection &&
            !selectionModel->isSelected(proxyModel->index(i, 0)))
            continue;

        destinationArray.append("\n");
        for (int j = 0; j < proxyColumnCount; ++j)
        {
            QVariant actualField = proxyModel->index(i, j).data();

            if (!actualField.isNull())
                variantToString(actualField, destinationArray, separator_);

            if (j != proxyColumnCount - 1)
                destinationArray.append(separator_);
        }
        //        bar.updateProgress(i + 1);
    }

    return destinationArray;
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
            QString value{variant.toString()};
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
