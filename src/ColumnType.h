#pragma once

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
