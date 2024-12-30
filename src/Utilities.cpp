#include "Utilities.h"

namespace utilities
{
QDate getStartOfExcelWorld()
{
    static const QDate startOfTheExcelWorld(1899, 12, 30);
    return startOfTheExcelWorld;
}

QHash<QByteArray, int> generateExcelColumnNames(int columnsNumber)
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

    QByteArray currentPrefix{QByteArrayLiteral("")};
    QHash<QByteArray, int> columnNamesWithIndexes;
    for (int i{0}; i < columnsNumber; ++i)
    {
        columnNamesWithIndexes[currentPrefix +
                               templateNames[i % templateNames.count()]] = i;

        if ((i != 0) && (((i + 1) % templateNames.count()) == 0))
            currentPrefix =
                templateNames[(i / (templateNames.count() - 1)) - 1];
    }
    return columnNamesWithIndexes;
}

int getMaxExcelColumns()
{
    constexpr int MAX_EXCEL_COLUMNS{600};
    return MAX_EXCEL_COLUMNS;
}

QString getXlsxTemplateName()
{
    QString templateFileName{QStringLiteral("template.xlsx")};
    return templateFileName;
}

QVariant getDefaultVariantForFormat(ColumnType format)
{
    switch (format)
    {
        case ColumnType::STRING:
            return QVariant(QMetaType(QMetaType::QString));

        case ColumnType::NUMBER:
            return QVariant(QMetaType(QMetaType::Double));

        case ColumnType::DATE:
            return QVariant(QMetaType(QMetaType::QDate));

        case ColumnType::UNKNOWN:
            Q_ASSERT(false);
    }
    return QVariant(QMetaType(QMetaType::QString));
}

}  // namespace utilities
