#ifndef EXPORTDSV_H
#define EXPORTDSV_H

#include "ExportData.h"

#include "eible_global.h"

#include <QLocale>

class QAbstractItemView;
class QIODevice;

class EIBLE_EXPORT ExportDsv : public ExportData
{
    Q_OBJECT
public:
    explicit ExportDsv(char separator);
    ~ExportDsv() override = default;

    ExportDsv& operator=(const ExportDsv& other) = delete;
    ExportDsv(const ExportDsv& other) = delete;

    ExportDsv& operator=(ExportDsv&& other) = delete;
    ExportDsv(ExportDsv&& other) = delete;

    void setDateFormat(Qt::DateFormat format);
    void setDateFormat(QString format);

    void setNumbersLocale(QLocale locale);

protected:
    bool writeContent(const QByteArray& content, QIODevice& ioDevice) override;

    QByteArray generateContent(const QAbstractItemView& view) override;

private:
    void variantToString(const QVariant& variant, QByteArray& destinationArray,
                         char separator) const;

    const char separator_;

    Qt::DateFormat qtDateFormat_{Qt::ISODate};
    QString dateFormat_{};

    QLocale locale_;
};

#endif  // EXPORTDSV_H
