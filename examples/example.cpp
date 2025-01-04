#include <iostream>

#include <QAbstractItemView>
#include <QApplication>
#include <QDate>
#include <QDebug>
#include <QFile>
#include <QTableWidget>
#include <QTimer>

#include <eible/ExportDsv.h>
#include <eible/ExportXlsx.h>
#include <eible/ImportOds.h>
#include <eible/ImportXlsx.h>

namespace
{
void initTable(QTableWidget& tableWidget)
{
    const int columnCount{3};
    const int rowCount{3};
    tableWidget.setRowCount(rowCount);
    tableWidget.setColumnCount(columnCount);
    tableWidget.setHorizontalHeaderLabels({QStringLiteral("Text"),
                                           QStringLiteral("Numeric"),
                                           QStringLiteral("Date")});
    for (int column = 0; column < columnCount; ++column)
    {
        for (int row = 0; row < rowCount; ++row)
        {
            auto* item = new QTableWidgetItem(
                "Item " + QString::number(column) + " " + QString::number(row));
            switch (column)
            {
                case 0:
                    item->setData(Qt::DisplayRole, item->text());
                    break;
                case 1:
                    item->setData(Qt::DisplayRole, 10 * column + row);
                    break;
                case 2:
                    item->setData(Qt::DisplayRole,
                                  QDate(2020, 1 + column, 1 + row));
                    break;
                default:
                    break;
            }
            tableWidget.setItem(row, column, item);
        }
    }
}

bool exportXlsx(const QTableWidget& tableWidget)
{
    const QString file{QCoreApplication::applicationDirPath() +
                       "/ExportedFile.xlsx"};
    ExportXlsx exportXlsx;
    QObject::connect(
        &exportXlsx, &ExportData::progressPercentChanged, &tableWidget,
        [](int progress)
        { std::cout << "Progress: " << progress << "%." << std::endl; });

    std::cout << "Exporting XLSX to " << file.toStdString() << "." << std::endl;
    QFile outFile(file);
    outFile.open(QIODevice::WriteOnly);
    const bool success = exportXlsx.exportView(tableWidget, outFile);
    if (success)
        std::cout << "Exporting XLSX successful." << std::endl;
    else
        std::cout << "Exporting XLSX failed." << std::endl;

    return success;
}

bool exportCsv(const QTableWidget& tableWidget)
{
    const QString file{QCoreApplication::applicationDirPath() +
                       "/ExportedFile.csv"};
    ExportDsv exportDsv(',');
    QObject::connect(
        &exportDsv, &ExportData::progressPercentChanged, &tableWidget,
        [](int progress)
        { std::cout << "Progress: " << progress << "%." << std::endl; });

    std::cout << "Exporting CSV to " << file.toStdString() << "." << std::endl;
    QFile outFile(file);
    outFile.open(QIODevice::WriteOnly);
    const bool success = exportDsv.exportView(tableWidget, outFile);
    if (success)
        std::cout << "Exporting CSV successful." << std::endl;
    else
        std::cout << "Exporting CSV failed." << std::endl;

    return success;
}

bool exportFiles()
{
    QTableWidget tableWidget;
    initTable(tableWidget);
    bool success = exportXlsx(tableWidget);
    success &= exportCsv(tableWidget);
    return success;
}

void printSpreadsheetContent(const QStringList& columnNames,
                             const QVector<QVector<QVariant>>& data)
{
    std::cout << columnNames.join('\t').toStdString() << std::endl;
    for (const auto& row : data)
    {
        for (const auto& column : row)
            std::cout << column.toString().toStdString() << "\t";
        std::cout << std::endl;
    }
}

bool importFile(ImportSpreadsheet& importer, const QString& fileName)
{
    const QString sheetName(QStringLiteral("Sheet1"));
    auto [sucess, columnNames] = importer.getColumnNames(sheetName);
    if (!sucess)
        return false;
    QVector<QVector<QVariant>> xlsxData;
    std::tie(sucess, xlsxData) = importer.getData(sheetName, {});
    if (!sucess)
        return false;
    std::cout << "File " + fileName.toStdString() + " opened, content of " +
                     sheetName.toStdString() + ":"
              << std::endl;
    printSpreadsheetContent(columnNames, xlsxData);
    return true;
}

bool importFiles()
{
    QFile odsFile(QStringLiteral(":/example.ods"));
    ImportOds importOds(odsFile);
    if (!importFile(importOds, odsFile.fileName()))
        return false;

    QFile xlsxFile(QStringLiteral(":/example.xlsx"));
    ImportXlsx importXlsx(xlsxFile);
    if (!importFile(importXlsx, xlsxFile.fileName()))
        return false;
    const QStringList sharedStrings{importXlsx.getSharedStrings().second};
    std::cout << "Where shared strings: ";
    for (int i = 0; i < sharedStrings.size(); ++i)
        std::cout << "(" << i << ", " << sharedStrings[i].toStdString() << ") ";
    std::cout << std::endl;
    return true;
}

bool performOperations() { return importFiles() && exportFiles(); }
}  // namespace

int main(int argc, char* argv[])
{
    const QApplication a(argc, argv);
    QTimer::singleShot(0, &a,
                       []() {
                           QApplication::exit(performOperations()
                                                  ? EXIT_SUCCESS
                                                  : EXIT_FAILURE);
                       });
    return QApplication::exec();
}
