#include "ExportXlsx.h"

#include <quazip5/quazipfile.h>

#include <QAbstractItemView>
#include <QCoreApplication>
#include <QFile>
#include <QVariant>

#include "EibleUtilities.h"

bool ExportXlsx::exportView(const QAbstractItemView& view, QIODevice& ioDevice)
{
    QFile xlsxTemplate(QStringLiteral(":/") +
                       EibleUtilities::getXlsxTemplateName());
    QuaZip inZip(&xlsxTemplate);
    inZip.open(QuaZip::mdUnzip);

    // Files list in template.
    QStringList fileList = inZip.getFileNameList();

    // Create out xlsx.
    QuaZip outZip(&ioDevice);
    outZip.open(QuaZip::mdCreate);

    // For each file in template.
    for (const QString& file : fileList)
    {
        inZip.setCurrentFile(file);

        QuaZipFile inZipFile(&inZip);
        if (!inZipFile.open(QIODevice::ReadOnly | QIODevice::Text))
            return false;

        QuaZipFile outZipFile(&outZip);
        if (!outZipFile.open(QIODevice::WriteOnly, QuaZipNewInfo(file)))
            return false;

        // Modify sheet1 by inserting data from view, copy rest of files.
        if (file.endsWith(QLatin1String("sheet1.xml")))
        {
            const QByteArray rowsContent = gatherSheetContent(view);

            // Replace empty tag by gathered data.
            QByteArray content = inZipFile.readAll();
            content.replace("<sheetData/>", rowsContent);

            outZipFile.write(content);
        }
        else
        {
            outZipFile.write(inZipFile.readAll());
        }

        inZipFile.close();
        outZipFile.close();
    }

    outZip.close();

    return true;
}

const QByteArray& ExportXlsx::getCellTypeTag(QVariant& cell)
{
    switch (cell.type())
    {
        case QVariant::Date:
        case QVariant::DateTime:
        {
            static const QByteArray dateTag{"s=\"3\""};
            cell = QVariant(-1 * cell.toDate().daysTo(
                                     EibleUtilities::getStartOfExcelWorld()));
            return dateTag;
        }

        case QVariant::Int:
        case QVariant::Double:
        {
            static const QByteArray numericTag{"s=\"4\""};
            return numericTag;
        }

        default:
        {
            static const QByteArray stringTag{"t=\"str\""};
            return stringTag;
        }
    }
}

void ExportXlsx::addHeaders(QByteArray& rowsContent,
                            const QAbstractItemModel& proxyModel,
                            const QStringList& columnNames) const
{
    rowsContent.append(R"(<row r="1" spans="1:1" x14ac:dyDescent="0.25">)");
    for (int j = 0; j < proxyModel.columnCount(); ++j)
    {
        QString header = proxyModel.headerData(j, Qt::Horizontal).toString();
        QString clearedHeader(
            header
                .replace(QRegExp(QStringLiteral("[<>&\"']")),
                         QStringLiteral(" "))
                .replace(QStringLiteral("\r\n"), QStringLiteral(" ")));
        rowsContent.append("<c r=\"" + columnNames.at(j));
        rowsContent.append(R"(1" t="str" s="6"><v>)" + clearedHeader +
                           "</v></c>");
    }
    rowsContent.append("</row>");
}

QByteArray ExportXlsx::gatherSheetContent(const QAbstractItemView& view)
{
    const auto proxyModel = view.model();
    Q_ASSERT(proxyModel != nullptr);

    const int proxyColumnCount = proxyModel->columnCount();
    if (proxyColumnCount == 0)
        return QByteArrayLiteral("</sheetData>");

    const QStringList columnNames =
        EibleUtilities::generateExcelColumnNames(proxyColumnCount);

    QByteArray rowsContent{QByteArrayLiteral("<sheetData>")};
    addHeaders(rowsContent, *proxyModel, columnNames);

    const bool multiSelection =
        (QAbstractItemView::MultiSelection == view.selectionMode());
    const QItemSelectionModel* selectionModel = view.selectionModel();

    const int proxyRowCount = proxyModel->rowCount();
    int skippedRows = 0;

    // For each row.
    for (int i = 0; i < proxyRowCount; ++i)
    {
        if (multiSelection &&
            !selectionModel->isSelected(proxyModel->index(i, 0)))
        {
            skippedRows++;
            continue;
        }

        // Create new row.
        const QByteArray rowNumber{QByteArray::number(i - skippedRows + 2)};
        rowsContent.append(QByteArrayLiteral("<row r=\""));
        rowsContent.append(rowNumber);
        rowsContent.append(R"(" spans="1:1" x14ac:dyDescent="0.25">)");

        // For each column.
        for (int j = 0; j < proxyColumnCount; ++j)
        {
            QVariant cell = proxyModel->index(i, j).data();
            if (cell.isNull())
                continue;

            // Create cell.
            rowsContent.append(QByteArrayLiteral("<c r=\""));
            rowsContent.append(columnNames.at(j));
            rowsContent.append(rowNumber);
            rowsContent.append(QByteArrayLiteral("\" "));
            rowsContent.append(getCellTypeTag(cell));
            rowsContent.append(QByteArrayLiteral("><v>"));
            rowsContent.append(cell.toByteArray());
            rowsContent.append(QByteArrayLiteral("</v></c>"));
        }
        rowsContent.append(QByteArrayLiteral("</row>"));

        Q_EMIT updateProgress(i + 1);
        QCoreApplication::processEvents();
    }

    // Close XML tag for data.
    rowsContent.append(QByteArrayLiteral("</sheetData>"));

    return rowsContent;
}
