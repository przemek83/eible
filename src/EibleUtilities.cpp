#include "EibleUtilities.h"

namespace EibleUtilities
{
QDate getStartOfExcelWorld()
{
    static const QDate startOfTheExcelWorld(1899, 12, 30);
    return startOfTheExcelWorld;
}

QList<QByteArray> generateExcelColumnNames(int columnsNumber)
{
    QList<QByteArray> templateNames;
    templateNames << QByteArrayLiteral("A") << QByteArrayLiteral("B")
                  << QByteArrayLiteral("C") << QByteArrayLiteral("D")
                  << QByteArrayLiteral("E") << QByteArrayLiteral("F")
                  << QByteArrayLiteral("G") << QByteArrayLiteral("H")
                  << QByteArrayLiteral("I") << QByteArrayLiteral("J")
                  << QByteArrayLiteral("K") << QByteArrayLiteral("L")
                  << QByteArrayLiteral("M") << QByteArrayLiteral("N")
                  << QByteArrayLiteral("O") << QByteArrayLiteral("P")
                  << QByteArrayLiteral("Q") << QByteArrayLiteral("R")
                  << QByteArrayLiteral("S") << QByteArrayLiteral("T")
                  << QByteArrayLiteral("U") << QByteArrayLiteral("V")
                  << QByteArrayLiteral("W") << QByteArrayLiteral("X")
                  << QByteArrayLiteral("Y") << QByteArrayLiteral("Z");

    QByteArray currentPrefix(QByteArrayLiteral(""));
    QList<QByteArray> columnNames;
    columnNames.reserve(columnsNumber);
    for (int i = 0; i < columnsNumber; ++i)
    {
        columnNames.append(currentPrefix +
                           templateNames[i % templateNames.count()]);

        if (i != 0 && (i + 1) % templateNames.count() == 0)
            currentPrefix = templateNames[i / (templateNames.count() - 1) - 1];
    }

    return columnNames;
}

int getMaxExcelColumns()
{
    constexpr int MAX_EXCEL_COLUMNS{600};
    return MAX_EXCEL_COLUMNS;
}

QString getXlsxTemplateName()
{
    const QString templateFileName{QStringLiteral("template.xlsx")};
    return templateFileName;
}

QVariant getDefaultVariantForFormat(ColumnType format)
{
    switch (format)
    {
        case ColumnType::STRING:
            return QVariant(QVariant::String);

        case ColumnType::NUMBER:
            return QVariant(QVariant::Double);

        case ColumnType::DATE:
            return QVariant(QVariant::Date);

        case ColumnType::UNKNOWN:
            Q_ASSERT(false);
    }
    return QVariant(QVariant::String);
}

}  // namespace EibleUtilities
