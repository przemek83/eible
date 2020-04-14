#ifndef EXPORTDSV_H
#define EXPORTDSV_H

#include "Export.h"

#include <QLocale>

class QAbstractItemView;
class QIODevice;

class ExportDsv : public Export
{
    Q_OBJECT
public:
    explicit ExportDsv(char separator);
    ~ExportDsv() override = default;

    ExportDsv& operator=(const ExportDsv& other) = delete;
    ExportDsv(const ExportDsv& other) = delete;

    ExportDsv& operator=(ExportDsv&& other) = delete;
    ExportDsv(ExportDsv&& other) = delete;

    bool exportView(const QAbstractItemView& view,
                    QIODevice& ioDevice) override;

    void setDateFormat(Qt::DateFormat format);
    void setDateFormat(QString format);

    void setLocale(QLocale locale);

private:
    void variantToString(const QVariant& variant, QByteArray& destinationArray,
                         char separator) const;

    const char separator_;

    Qt::DateFormat qtDateFormat_{Qt::ISODate};
    QString dateFormat_{};

    QLocale locale_;
};

#endif  // EXPORTDSV_H
