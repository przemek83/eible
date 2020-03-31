#ifndef TOOLS_H
#define TOOLS_H

#include <QDate>

namespace Utilities
{
const QDate& getStartOfExcelWorld();

QStringList generateExcelColumnNames(int columnsNumber);

static constexpr int MAX_EXCEL_COLUMNS {600};
}

#endif // TOOLS_H
