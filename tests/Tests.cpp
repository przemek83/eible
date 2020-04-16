#include <QTest>

#include "ExportDsvTest.h"
#include "ExportXlsxTest.h"
#include "ImportXlsxTest.h"

int main(int argc, char* argv[])
{
    QLocale::setDefault(QLocale::c());
    QApplication a(argc, argv);

    ExportXlsxTest exportXlsxTest;
    QTest::qExec(&exportXlsxTest);

    ExportDsvTest exportDsvTest;
    QTest::qExec(&exportDsvTest);

    ImportXlsxTest importXlsxTest;
    QTest::qExec(&importXlsxTest);

    return 0;
}
