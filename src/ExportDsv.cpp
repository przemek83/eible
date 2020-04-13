#include "ExportDsv.h"

#include <QAbstractItemView>
#include <QDate>
#include <QVariant>

ExportDsv::ExportDsv(char separator) : Export(), separator_(separator) {}

bool ExportDsv::exportView(const QAbstractItemView& view, QIODevice& ioDevice)
{
    /*
     * Using standards described on:
     * http://en.wikipedia.org/wiki/Comma-separated_values
     * and
     * http://tools.ietf.org/html/rfc4180
     */

    const auto proxyModel = view.model();
    Q_ASSERT(proxyModel != nullptr);

    int proxyColumnCount = proxyModel->columnCount();
    if (proxyColumnCount == 0)
        return true;

    QByteArray destinationArray;
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

    return ioDevice.write(destinationArray) != -1;
}

QString doubleToStringUsingLocale(double value, int precision)
{
    static bool initialized{false};
    static QLocale locale = QLocale::system();
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
            QString value(doubleToStringUsingLocale(variant.toDouble(), 2));
            destinationArray.append(value);
            break;
        }

        case QVariant::Date:
        case QVariant::DateTime:
        {
            destinationArray.append(variant.toDate().toString(Qt::ISODate));
            break;
        }

        case QVariant::String:
        {
            QString value{variant.toString()};
            value.replace('"', QStringLiteral("\"\""));
            if (value.contains(separator))
                destinationArray.append("\"" + value + "\"");
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
