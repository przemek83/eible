#pragma once

#include <QDate>
#include <QVariant>

#include <eible/ColumnType.h>

namespace utilities
{
QDate getStartOfExcelWorld();

QHash<QByteArray, int> generateExcelColumnNames(int columnsNumber);

int getMaxExcelColumns();

QString getXlsxTemplateName();

QVariant getDefaultVariantForFormat(ColumnType format);
}  // namespace utilities
