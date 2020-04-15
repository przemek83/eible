#include <iostream>

#include <QAbstractItemView>
#include <QApplication>
#include <QDate>
#include <QDebug>
#include <QFile>
#include <QTableWidget>
#include <QTimer>

#include "ExportDsv.h"
#include "ExportXlsx.h"

void initTable(QTableWidget& tableWidget)
{
    const int columnCount{3};
    const int rowCount{3};
    tableWidget.setRowCount(rowCount);
    tableWidget.setColumnCount(columnCount);
    tableWidget.setHorizontalHeaderLabels({"Text", "Numeric", "Date"});
    for (int column = 0; column < columnCount; ++column)
        for (int row = 0; row < rowCount; ++row)
        {
            auto item = new QTableWidgetItem("Item " + QString::number(column) +
                                             " " + QString::number(row));
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
            }
            tableWidget.setItem(row, column, item);
        }
}

static bool exportXlsx(const QTableWidget& tableWidget)
{
    QString file{QCoreApplication::applicationDirPath() + "/ExportedFile.xlsx"};
    ExportXlsx exportXlsx;
    QObject::connect(&exportXlsx, &ExportData::updateProgress, &tableWidget,
                     [&tableWidget](int progress) {
                         std::cout << "Progress: " << progress << "/"
                                   << tableWidget.rowCount() << "."
                                   << std::endl;
                     });

    std::cout << "Exporting XLSX to " << file.toStdString() << "." << std::endl;
    QFile outFile(file);
    outFile.open(QIODevice::WriteOnly);
    bool success = exportXlsx.exportView(tableWidget, outFile);
    if (success)
        std::cout << "Exporting XLSX successful." << std::endl;
    else
        std::cout << "Exporting XLSX failed." << std::endl;

    return success;
}

static bool exportCsv(const QTableWidget& tableWidget)
{
    QString file{QCoreApplication::applicationDirPath() + "/ExportedFile.csv"};
    ExportDsv exportDsv(',');
    QObject::connect(&exportDsv, &ExportData::updateProgress, &tableWidget,
                     [&tableWidget](int progress) {
                         std::cout << "Progress: " << progress << "/"
                                   << tableWidget.rowCount() << "."
                                   << std::endl;
                     });

    std::cout << "Exporting CSV to " << file.toStdString() << "." << std::endl;
    QFile outFile(file);
    outFile.open(QIODevice::WriteOnly);
    bool success = exportDsv.exportView(tableWidget, outFile);
    if (success)
        std::cout << "Exporting CSV successful." << std::endl;
    else
        std::cout << "Exporting CSV failed." << std::endl;

    return success;
}

static bool exportFiles()
{
    QTableWidget tableWidget;
    initTable(tableWidget);

    bool success = exportXlsx(tableWidget);
    success &= exportCsv(tableWidget);
    return success;
}

int main(int argc, char* argv[])
{
    QApplication a(argc, argv);
    QTimer::singleShot(0, []() {
        QApplication::exit(exportFiles() ? EXIT_SUCCESS : EXIT_FAILURE);
    });
    return a.exec();
}
