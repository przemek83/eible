#pragma once

#include <QtCore/qglobal.h>

#if defined(EIBLE_LIBRARY)
#define EIBLE_EXPORT Q_DECL_EXPORT
#else
#define EIBLE_EXPORT Q_DECL_IMPORT
#endif
