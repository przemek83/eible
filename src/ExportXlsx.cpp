#include <eible/ExportXlsx.h>

#include <quazip/quazipfile.h>
#include <QAbstractItemView>
#include <QCoreApplication>
#include <QFile>
#include <QRegExp>
#include <QVariant>

#include "Utilities.h"

ExportXlsx::ExportXlsx(QObject* parent) : ExportData(parent) {}

bool ExportXlsx::writeContent(const QByteArray& content, QIODevice& ioDevice)
{
    QFile xlsxTemplate(QStringLiteral(":/") + utilities::getXlsxTemplateName());
    QuaZip inZip(&xlsxTemplate);
    inZip.open(QuaZip::mdUnzip);

    // Files list in template.
    const QStringList fileList{inZip.getFileNameList()};

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
        if (file.endsWith(QString::fromLatin1("sheet1.xml")))
        {
            // Replace empty tag by gathered data.
            QByteArray contentFile{inZipFile.readAll()};
            contentFile.replace("<sheetData/>", content);

            outZipFile.write(contentFile);
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

QByteArray ExportXlsx::getEmptyContent()
{
    return QByteArrayLiteral("</sheetData>");
}

QByteArray ExportXlsx::generateHeaderContent(const QAbstractItemModel& model)
{
    const int columnCount{model.columnCount()};
    if (columnNames_.size() != columnCount)
        initColumnNames(columnCount);

    QByteArray headersContent{QByteArrayLiteral("<sheetData>")};
    headersContent.append(R"(<row r="1" spans="1:1" x14ac:dyDescent="0.25">)");

    for (int j{0}; j < columnCount; ++j)
    {
        QString header{model.headerData(j, Qt::Horizontal).toString()};
        const QString clearedHeader(
            header
                .replace(QRegularExpression(QStringLiteral("[<>&\"']")),
                         QStringLiteral(" "))
                .replace(QStringLiteral("\r\n"), QStringLiteral(" ")));
        headersContent.append("<c r=\"" + columnNames_[j]);
        headersContent.append(R"(1" t="str" s="6"><v>)" +
                              clearedHeader.toUtf8() + "</v></c>");
    }
    headersContent.append("</row>");
    return headersContent;
}

QByteArray ExportXlsx::generateRowContent(const QAbstractItemModel& model,
                                          int row, int skippedRowsCount)
{
    const int columnCount{model.columnCount()};
    if (columnNames_.size() != columnCount)
        initColumnNames(columnCount);

    QByteArray rowContent;
    const QByteArray rowNumber{
        QByteArray::number((row - skippedRowsCount) + 2)};
    rowContent.append(QByteArrayLiteral("<row r=\""));
    rowContent.append(rowNumber);
    rowContent.append(R"(" spans="1:1" x14ac:dyDescent="0.25">)");
    for (int column{0}; column < columnCount; ++column)
    {
        const QVariant& cell{model.index(row, column).data()};
        if (!cell.isNull())
        {
            rowContent.append(CELL_START);
            rowContent.append(columnNames_[column]);
            rowContent.append(rowNumber);
            rowContent.append(ROW_NUMBER_CLOSE);
            rowContent.append(getCellTypeTag(cell));
            rowContent.append(VALUE_START);
            if ((cell.typeId() == QMetaType::QDate) ||
                (cell.typeId() == QMetaType::QDateTime))
                rowContent.append(QByteArray::number(
                    -1 *
                    cell.toDate().daysTo(utilities::getStartOfExcelWorld())));
            else
                rowContent.append(cell.toByteArray());
            rowContent.append(CELL_END);
        }
    }
    rowContent.append(QByteArrayLiteral("</row>"));
    return rowContent;
}

QByteArray ExportXlsx::getContentEnding()
{
    return QByteArrayLiteral("</sheetData>");
}

const QByteArray& ExportXlsx::getCellTypeTag(const QVariant& cell) const
{
    switch (cell.typeId())
    {
        case QMetaType::QDate:
        case QMetaType::QDateTime:
            return DATE_TAG;

        case QMetaType::Int:
        case QMetaType::Double:
            return NUMERIC_TAG;

        default:
            return STRING_TAG;
    }
}

void ExportXlsx::initColumnNames(int modelColumnCount)
{
    columnNames_.clear();
    QHash<QByteArray, int> nameToIndexMap{
        utilities::generateExcelColumnNames(modelColumnCount)};
    columnNames_.resize(nameToIndexMap.size());

    const auto mapEnd{nameToIndexMap.end()};
    for (auto it{nameToIndexMap.begin()}; it != mapEnd; ++it)
        columnNames_[it.value()] = it.key();
}
