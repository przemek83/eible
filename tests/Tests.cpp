#include <QTest>

#include "ExportDsvTest.h"
#include "ExportXlsxTest.h"
#include "ImportOdsTest.h"
#include "ImportXlsxTest.h"

int main(int argc, char* argv[])
{
    QLocale::setDefault(QLocale::c());
    const QApplication a(argc, argv);

    ExportXlsxTest exportXlsxTest;
    QTest::qExec(&exportXlsxTest);

    ExportDsvTest exportDsvTest;
    QTest::qExec(&exportDsvTest);

    ImportXlsxTest importXlsxTest;
    QTest::qExec(&importXlsxTest);

    ImportOdsTest importOdsTest;
    QTest::qExec(&importOdsTest);

    return 0;
}
