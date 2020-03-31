#include "ExportXlsx.h"

#include <QAbstractItemView>
#include <QFile>
#include <quazip5/quazipfile.h>
#include <QVariant>

#include "Utilities.h"

ExportXlsx::ExportXlsx(const QString& filePath) : filePath_(filePath)
{

}

bool ExportXlsx::exportView(const QAbstractItemView* view)
{
    Q_ASSERT(view != nullptr);

    // QuaZip has problems with opening file from resources. Workaround it...
    QFile::copy(QStringLiteral(":/template.xlsx"), QStringLiteral("template.xlsx"));

    //Open xlsx template.
    QuaZip inZip(QStringLiteral("template.xlsx"));
    inZip.open(QuaZip::mdUnzip);

    //Files list in template.
    QStringList fileList = inZip.getFileNameList();

    //Create out xlsx.
    QuaZip outZip(filePath_);
    outZip.open(QuaZip::mdCreate);

    //For each file in template.
    for (const QString& file : fileList)
    {
        inZip.setCurrentFile(file);

        QuaZipFile inZipFile(&inZip);
        if (!inZipFile.open(QIODevice::ReadOnly | QIODevice::Text))
        {
//            LOG(LogTypes::IMPORT_EXPORT,
//                "Can not open file " + inZipFile.getFileName() + ".");
            return false;
        }

        QuaZipFile outZipFile(&outZip);
        if (!outZipFile.open(QIODevice::WriteOnly, QuaZipNewInfo(file)))
        {
//            LOG(LogTypes::IMPORT_EXPORT, "Can not open file " + inZipFile.getFileName() + ".");
            return false;
        }

        //Modify sheet1 by inserting data from view, copy rest of files.
        if (file.endsWith(QLatin1String("sheet1.xml")))
        {
            const QByteArray rowsContent = gatherSheetContent(view);

            //Replace empty tag by gathered data.
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

const QString& ExportXlsx::getCellTypeTag(QVariant& cell)
{
    switch (cell.type())
    {
        case QVariant::Date:
        case QVariant::DateTime:
        {
            static const QString dateTag(QStringLiteral("s=\"3\""));
            cell = QVariant(-1 * cell.toDate().daysTo(Utilities::getStartOfExcelWorld()));
            return dateTag;
        }

        case QVariant::Int:
        case QVariant::Double:
        {
            static const QString numericTag(QStringLiteral("s=\"4\""));
            return numericTag;
        }

        default:
        {
            static const QString stringTag{QStringLiteral("t=\"str\"")};
            return stringTag;
        }
    }
}

QByteArray ExportXlsx::gatherSheetContent(const QAbstractItemView* view)
{
    QByteArray rowsContent;
    rowsContent.append("<sheetData>");

    QStringList columnNames = Utilities::generateExcelColumnNames(Utilities::MAX_EXCEL_COLUMNS);
    auto proxyModel = qobject_cast<QAbstractItemModel*>(view->model());

    Q_ASSERT(nullptr != proxyModel);

    bool multiSelection =
        (QAbstractItemView::MultiSelection == view->selectionMode());
    QItemSelectionModel* selectionModel = view->selectionModel();

    int proxyColumnCount = proxyModel->columnCount();
    int proxyRowCount = proxyModel->rowCount();

    //Add headers using string type.
    rowsContent.append(R"(<row r="1" spans="1:1" x14ac:dyDescent="0.25">)");
    for (int j = 0; j < proxyColumnCount; ++j)
    {
        QString header = proxyModel->headerData(j, Qt::Horizontal).toString();
        QString clearedHeader(header.replace(QRegExp(QStringLiteral("[<>&\"']")), QStringLiteral(" ")).replace(QStringLiteral("\r\n"), QStringLiteral(" ")));
        rowsContent.append("<c r=\"" + columnNames.at(j));
        rowsContent.append(R"(1" t="str" s="6"><v>)" + clearedHeader + "</v></c>");
    }
    rowsContent.append("</row>");

    int skippedRows = 0;

//    const QString barTitle =
//        Constants::getProgressBarTitle(Constants::BarTitle::SAVING);
//    ProgressBarCounter bar(barTitle, proxyModel->rowCount(), nullptr);
//    bar.showDetached();

    //For each row.
    for (int i = 0; i < proxyRowCount; ++i)
    {
        if (multiSelection &&
            !selectionModel->isSelected(proxyModel->index(i, 0)))
        {
            skippedRows++;
            continue;
        }

        //Create new row.
        QString rowNumber(QString::fromLatin1(QByteArray::number(i - skippedRows + 2)));
        rowsContent.append("<row r=\"" + rowNumber);
        rowsContent.append(R"(" spans="1:1" x14ac:dyDescent="0.25">)");

        //For each column.
        for (int j = 0; j < proxyColumnCount; ++j)
        {
            QVariant cell = proxyModel->index(i, j).data();
            if (cell.isNull())
                continue;

            //Create cell.
            rowsContent.append("<c r=\"" + columnNames.at(j) + rowNumber + "\" ");
            rowsContent.append(getCellTypeTag(cell));
            rowsContent.append("><v>" + cell.toByteArray() + "</v></c>");
        }
        rowsContent.append("</row>");

        Q_EMIT updateProgress(i + 1);
    }

    //Close XML tag for data.
    rowsContent.append("</sheetData>");

    return rowsContent;
}
