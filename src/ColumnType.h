#ifndef COLUMNTYPE_H
#define COLUMNTYPE_H

#include "eible_global.h"

enum class EIBLE_EXPORT ColumnType : char
{
    UNKNOWN = -1,
    STRING,
    NUMBER,
    DATE
};

#endif  // COLUMNTYPE_H
