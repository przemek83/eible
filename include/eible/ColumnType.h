#pragma once

#include <QMetaType>

enum class ColumnType : char
{
    UNKNOWN = -1,
    STRING,
    NUMBER,
    DATE
};

Q_DECLARE_METATYPE(ColumnType)
