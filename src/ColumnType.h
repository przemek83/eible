#ifndef COLUMNTYPE_H
#define COLUMNTYPE_H

#include <QMetaType>

#include "eible_global.h"

enum class ColumnType : char
{
    UNKNOWN = -1,
    STRING,
    NUMBER,
    DATE
};

Q_DECLARE_METATYPE(ColumnType)

#endif  // COLUMNTYPE_H
