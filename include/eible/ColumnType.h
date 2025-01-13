#pragma once

#include <QMetaType>

enum class ColumnType : signed char
{
    UNKNOWN = -1,
    STRING,
    NUMBER,
    DATE
};

Q_DECLARE_METATYPE(ColumnType)
