#include <iostream>

#include <QAbstractItemView>
#include <QApplication>
#include <QDate>
#include <QDebug>
#include <QTableWidget>
#include <QTimer>

#include "ExportXlsx.h"

void initTable(QTableWidget& tableWidget)
{
    const int columnCount {3};
    const int rowCount {3};
    tableWidget.setRowCount(rowCount);
    tableWidget.setColumnCount(columnCount);
    tableWidget.setHorizontalHeaderLabels({"Text column",
                                           "Numeric column",
                                           "Date column"});
    for (int column = 0; column < columnCount; ++column)
        for (int row = 0; row < rowCount; ++row)
        {
            auto item = new QTableWidgetItem("Item " + QString::number(column) +
                                             "," + QString::number(row));
            switch (column)
            {
                case 0:
                    item->setData(Qt::DisplayRole, item->text());
                    break;
                case 1:
                    item->setData(Qt::DisplayRole, 10 * column + row);
                    break;
                case 2:
                    item->setData(Qt::DisplayRole, QDate(2020, 1 + column, 1 + row));
                    break;
            }
            tableWidget.setItem(row, column, item);
        }
}

static bool exportFile()
{
    QTableWidget tableWidget;
    initTable(tableWidget);

    QString file {QCoreApplication::applicationDirPath() + "/ExportedFile.xlsx"};
    ExportXlsx exportXlsx(file);
    QObject::connect(&exportXlsx, &ExportXlsx::updateProgress, &exportXlsx,
                     [&tableWidget](int progress)
    {
        std::cout << "Progress: " << progress << "/" << tableWidget.rowCount()
                  << "." << std::endl;
    });

    std::cout << "Exporting " << file.toStdString() << "." << std::endl;
    bool success = exportXlsx.exportView(&tableWidget);
    if (success)
        std::cout << "Exporting successful." << std::endl;
    else
        std::cout << "Exporting failed." << std::endl;

    return success;
}

int main(int argc, char* argv[])
{
    QApplication a(argc, argv);
    QTimer::singleShot(0, []()
    {
        QApplication::exit(exportFile() ? EXIT_SUCCESS : EXIT_FAILURE);
    });
    return a.exec();
}
