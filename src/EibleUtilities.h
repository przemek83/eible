#ifndef TOOLS_H
#define TOOLS_H

#include <QDate>

#include "eible_global.h"

namespace EibleUtilities
{
const QDate& EIBLE_EXPORT getStartOfExcelWorld();

QStringList EIBLE_EXPORT generateExcelColumnNames(int columnsNumber);

int EIBLE_EXPORT getMaxExcelColumns();

QString EIBLE_EXPORT getXlsxTemplateName();
}  // namespace EibleUtilities

#endif  // TOOLS_H
