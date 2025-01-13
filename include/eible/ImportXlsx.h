#pragma once

#include <optional>

#include <QHash>
#include <QtXml/QDomDocument>

#include "ImportSpreadsheet.h"
#include "eible_global.h"

class QuaZip;
class QuaZipFile;
class QXmlStreamReader;

/// @class ImportXlsx
/// @brief Class analysing content and loading it from xlsx files.
class EIBLE_EXPORT ImportXlsx : public ImportSpreadsheet
{
    Q_OBJECT
public:
    /// @brief Constructor.
    /// @param ioDevice Source of data (QFile, QBuffer, ...).
    explicit ImportXlsx(QIODevice& ioDevice);

    std::pair<bool, QStringList> getSheetNames() override;

    std::pair<bool, QVector<ColumnType>> getColumnTypes(
        const QString& sheetName) override;

    std::pair<bool, QStringList> getColumnNames(
        const QString& sheetName) override;

    /// @brief Get shared strings from xlsx.
    /// @return First value indicating success, second list of shared strings.
    std::pair<bool, QStringList> getSharedStrings();

    /// @brief Get date styles from xlsx.
    /// @return First value indicating success, second list of date style ids.
    std::pair<bool, QList<int>> getDateStyles();

    /// @brief Get all styles from xlsx.
    /// @return First value indicating success, second list of all style ids.
    std::pair<bool, QList<int>> getAllStyles();

    std::pair<bool, QVector<QVector<QVariant>>> getLimitedData(
        const QString& sheetName, const QVector<int>& excludedColumns,
        int rowLimit) override;

private:
    std::tuple<bool, std::optional<QList<int>>, std::optional<QList<int>>>
    getStyles();

    std::pair<bool, QString> getSheetPath(const QString& sheetName);

    void moveToSecondRow(QuaZipFile& quaZipFile,
                         QXmlStreamReader& reader) const;

    std::pair<bool, int> getCount(const QString& sheetName,
                                  const QHash<QString, int>& countMap) override;

    std::pair<bool, QMap<QString, QString>> getSheetIdToUserFriendlyNameMap();

    std::pair<bool, QVector<std::pair<QString, QString>>> retrieveSheets(
        const QMap<QString, QString>& sheetIdToUserFriendlyNameMap);

    template <typename NodesRetriever>
    std::pair<bool, QDomNodeList> getSheetNodes(
        QuaZipFile& quaZipFile, const NodesRetriever& nodesRetriever);

    bool isRowStart(const QXmlStreamReader& reader) const;
    bool isCellStart(const QXmlStreamReader& reader) const;
    bool isCellEnd(const QXmlStreamReader& reader) const;
    bool isVTagStart(const QXmlStreamReader& reader) const;

    std::pair<bool, QStringList> retrieveColumnNames(
        const QString& sheetName) override;

    std::tuple<bool, int, QVector<ColumnType>> retrieveRowCountAndColumnTypes(
        const QString& sheetName) override;

    static QList<int> retrieveDateStyles(const QDomNodeList& sheetNodes);
    static QList<int> retrieveAllStyles(const QDomNodeList& sheetNodes);

    QString getColumnName(QXmlStreamReader& reader,
                          const QString& currentColType) const;

    void skipToFirstRow(QXmlStreamReader& reader) const;

    int getCurrentColumnNumber(const QXmlStreamReader& reader,
                               int rowCountDigitsInXlsx) const;

    ColumnType recognizeColumnType(ColumnType currentType,
                                   const QXmlStreamReader& reader) const;

    QVariant getCurrentValueForStringColumn(const QString& currentColType,
                                            const QString& actualSTagValue,
                                            const QString& valueAsString) const;

    QVariant getCurrentValue(QXmlStreamReader& reader,
                             ColumnType currentColumnFormat,
                             const QString& currentColType,
                             const QString& actualSTagValue) const;

    static QDate getDateFromString(const QString& dateAsString);

    bool isDateStyle(const QString& sTagValue) const;

    bool isCommonDataOk();

    QStringList createSheetNames() const;

    static void adjustColumnTypeSize(QVector<ColumnType>& columnTypes,
                                     int column);

    std::optional<QVector<std::pair<QString, QString>>> sheets_{std::nullopt};
    std::optional<QStringList> sharedStrings_{std::nullopt};
    std::optional<QList<int>> dateStyles_{std::nullopt};
    std::optional<QList<int>> allStyles_{std::nullopt};
    const QHash<QByteArray, int> excelColNames_;

    static constexpr int decimalBase_{10};
    const QString rowTag_{QStringLiteral("row")};
    const QString cellTag_{QStringLiteral("c")};
    const QString sheetDataTag_{QStringLiteral("sheetData")};
    const QString sTag_{QStringLiteral("s")};
    const QString vTag_{QStringLiteral("v")};
    const QString rTag_{QStringLiteral("r")};
    const QString tTag_{QStringLiteral("t")};
    const QString strTag_{QStringLiteral("str")};
};
