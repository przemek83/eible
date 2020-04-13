#include <QTest>

#include "ExportDsvTest.h"
#include "ExportXlsxTest.h"

int main(int argc, char* argv[])
{
    QApplication a(argc, argv);

    ExportXlsxTest exportXlsxTest;
    QTest::qExec(&exportXlsxTest);

    ExportDsvTest exportDsvTest;
    QTest::qExec(&exportDsvTest);

    return 0;
}
