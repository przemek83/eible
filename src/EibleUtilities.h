#ifndef TOOLS_H
#define TOOLS_H

#include <QDate>
#include <QVariant>

#include "ColumnType.h"

namespace EibleUtilities
{
QDate getStartOfExcelWorld();

QList<QByteArray> generateExcelColumnNames(int columnsNumber);

int getMaxExcelColumns();

QString getXlsxTemplateName();

QVariant getDefaultVariantForFormat(ColumnType format);
}  // namespace EibleUtilities

#endif  // TOOLS_H
