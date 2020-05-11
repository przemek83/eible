#ifndef EXPORTDSV_H
#define EXPORTDSV_H

#include "ExportData.h"

#include "eible_global.h"

#include <QLocale>

class QAbstractItemView;
class QIODevice;

/**
 * @class ExportDsv
 * @brief Class for exporting data to DSV (Delimiter Separated Values) files.
 */
class EIBLE_EXPORT ExportDsv : public ExportData
{
    Q_OBJECT
public:
    /**
     * @brief Constructor.
     * @param separator Separator to be used during export (comma, tab, ...).
     */
    explicit ExportDsv(char separator);
    ~ExportDsv() override = default;

    ExportDsv& operator=(const ExportDsv& other) = delete;
    ExportDsv(const ExportDsv& other) = delete;

    ExportDsv& operator=(ExportDsv&& other) = delete;
    ExportDsv(ExportDsv&& other) = delete;

    /**
     * @brief Change date format to given one via Qt::DateFormat. (overloaded)
     * @param format New format for dates as Qt::DateFormat.
     */
    void setDateFormat(Qt::DateFormat format);

    /**
     * @brief Change date format to given one via string. (overloaded)
     * @param format New format for dates as string.
     */
    void setDateFormat(QString format);

    /**
     * @brief Set locale for numbers conversion.
     * @param locale Locale to be used in exporting.
     */
    void setNumbersLocale(QLocale locale);

protected:
    bool writeContent(const QByteArray& content, QIODevice& ioDevice) override;

    QByteArray getEmptyContent() override;

    QByteArray generateHeaderContent(const QAbstractItemModel& model) override;

    QByteArray generateRowContent(const QAbstractItemModel& model, int row,
                                  int skippedRowsCount) override;

private:
    void variantToString(const QVariant& variant, QByteArray& destinationArray,
                         char separator) const;

    const char separator_;

    Qt::DateFormat qtDateFormat_{Qt::ISODate};
    QString dateFormat_{};

    QLocale locale_;
};

#endif  // EXPORTDSV_H
