#ifndef TOOLS_H
#define TOOLS_H

#include <QDate>
#include <QVariant>

#include "ColumnType.h"
#include "eible_global.h"

namespace EibleUtilities
{
QDate EIBLE_EXPORT getStartOfExcelWorld();

QStringList EIBLE_EXPORT generateExcelColumnNames(int columnsNumber);

int EIBLE_EXPORT getMaxExcelColumns();

QString EIBLE_EXPORT getXlsxTemplateName();

QVariant getDefaultVariantForFormat(ColumnType format);
}  // namespace EibleUtilities

#endif  // TOOLS_H
