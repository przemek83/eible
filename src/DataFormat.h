#ifndef DATAFORMAT_H
#define DATAFORMAT_H

#include "eible_global.h"

enum class EIBLE_EXPORT DataFormat : char
{
    DATA_FORMAT_UNKNOWN = -1,
    DATA_FORMAT_STRING,
    DATA_FORMAT_FLOAT,
    DATA_FORMAT_DATE
};

#endif  // DATAFORMAT_H
