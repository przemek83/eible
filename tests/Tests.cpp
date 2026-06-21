#include <QTest>

#include "ExportDsvTest.h"
#include "ExportXlsxTest.h"
#include "ImportOdsTest.h"
#include "ImportXlsxTest.h"

int main(int argc, char* argv[])
{
    QLocale::setDefault(QLocale::c());
    const QApplication a(argc, argv);

    int status{EXIT_SUCCESS};
    ExportXlsxTest exportXlsxTest;
    status = std::max(status, QTest::qExec(&exportXlsxTest));

    ExportDsvTest exportDsvTest;
    status = std::max(status, QTest::qExec(&exportDsvTest));

    ImportXlsxTest importXlsxTest;
    status = std::max(status, QTest::qExec(&importXlsxTest));

    ImportOdsTest importOdsTest;
    status = std::max(status, QTest::qExec(&importOdsTest));

    return status;
}
