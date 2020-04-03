#include "EibleUtilities.h"

namespace EibleUtilities
{
const QDate& getStartOfExcelWorld()
{
    static const QDate startOfTheExcelWorld(1899, 12, 30);
    return startOfTheExcelWorld;
}

QStringList generateExcelColumnNames(int columnsNumber)
{
    static QStringList templateNames;

    if (templateNames.isEmpty())
    {
        templateNames << QStringLiteral("A") <<
                      QStringLiteral("B") << QStringLiteral("C") << QStringLiteral("D") <<
                      QStringLiteral("E") << QStringLiteral("F") << QStringLiteral("G") <<
                      QStringLiteral("H") << QStringLiteral("I") << QStringLiteral("J") <<
                      QStringLiteral("K") << QStringLiteral("L") << QStringLiteral("M") <<
                      QStringLiteral("N") << QStringLiteral("O") << QStringLiteral("P") <<
                      QStringLiteral("Q") << QStringLiteral("R") << QStringLiteral("S") <<
                      QStringLiteral("T") << QStringLiteral("U") << QStringLiteral("V") <<
                      QStringLiteral("W") << QStringLiteral("X") << QStringLiteral("Y") <<
                      QStringLiteral("Z");
    }

    QString currentPrefix(QLatin1String(""));
    QStringList columnNames;
    columnNames.reserve(columnsNumber);
    for (int i = 0; i < columnsNumber; ++i)
    {
        columnNames.append(currentPrefix + templateNames[i % templateNames.count()]);

        if (i != 0 && 0 == (i + 1) % templateNames.count())
        {
            currentPrefix = templateNames[i / (templateNames.count() - 1) - 1];
        }
    }

    return columnNames;
}

int getMaxExcelColumns()
{
    constexpr int MAX_EXCEL_COLUMNS {600};
    return MAX_EXCEL_COLUMNS;
}

QString getXlsxTemplateName()
{
    const QString templateFileName {"template.xlsx"};
    return templateFileName;
}


} // namespace Utilities
