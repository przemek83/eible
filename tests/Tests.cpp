#include <QTest>

#include "ExportXlsxTest.h"

int main(int argc, char* argv[])
{
    QApplication a(argc, argv);

    ExportXlsxTest exportXlsxTest;
    QTest::qExec(&exportXlsxTest);

    return 0;
}
