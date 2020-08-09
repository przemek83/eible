#pragma once

#include <QDate>
#include <QVariant>

#include "ColumnType.h"

namespace EibleUtilities
{
QDate getStartOfExcelWorld();

QHash<QByteArray, int> generateExcelColumnNames(int columnsNumber);

int getMaxExcelColumns();

QString getXlsxTemplateName();

QVariant getDefaultVariantForFormat(ColumnType format);
}  // namespace EibleUtilities
