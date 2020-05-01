#ifndef IMPORTODSTEST_H
#define IMPORTODSTEST_H

#include <QObject>

class ImportOdsTest : public QObject
{
    Q_OBJECT
private Q_SLOTS:
    void testRetrievingSheetNames();
    void testRetrievingSheetNamesFromEmptyFile();

    void testGetColumnCount_data();
    void testGetColumnCount();

    void testGetRowCount_data();
    void testGetRowCount();

private:
    static const QString testFileName_;
};

#endif  // IMPORTODSTEST_H
