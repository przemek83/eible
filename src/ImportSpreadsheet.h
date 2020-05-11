#ifndef IMPORTSPREADSHEET_H
#define IMPORTSPREADSHEET_H

#include <QObject>

#include <Qt5Quazip/quazip.h>

#include "ColumnType.h"
#include "eible_global.h"

class QIODevice;
class QuaZipFile;

/**
 * @class ImportSpreadsheet
 * @brief Base class for spreadsheet importing classes.
 */
class EIBLE_EXPORT ImportSpreadsheet : public QObject
{
    Q_OBJECT
public:
    /**
     * @brief Constructor.
     * @param ioDevice Source of data (QFile, QBuffer, ...).
     */
    explicit ImportSpreadsheet(QIODevice& ioDevice);

    /**
     * @brief Get list of sheet names.
     * @return First value indicating success, second is sheet names list.
     */
    virtual std::pair<bool, QStringList> getSheetNames() = 0;

    /**
     * @brief Get list of column names.
     * @param sheetName Sheet name.
     * @return First value indicating success, second is column names list.
     */
    virtual std::pair<bool, QStringList> getColumnNames(
        const QString& sheetName) = 0;

    /**
     * @brief Get list of column types.
     * @param sheetName Sheet name.
     * @return First value indicating success, second is column types vector.
     */
    virtual std::pair<bool, QVector<ColumnType>> getColumnTypes(
        const QString& sheetName) = 0;

    /**
     * @brief Get last error.
     * @return First value contains function name, second error.
     */
    std::pair<QString, QString> getError() const;

    /**
     * @brief Set name used for empty columns.
     * @param name Name for empty column.
     */
    void setNameForEmptyColumn(const QString& name);

    /**
     * @brief Get number of columns in given sheet.
     * @param sheetName Sheet name.
     * @return First value indicating success, second is column count.
     */
    virtual std::pair<bool, unsigned int> getColumnCount(
        const QString& sheetName) = 0;

    /**
     * @brief Get number of rows in given sheet.
     * @param sheetName Sheet name.
     * @return First value indicating success, second is row count.
     */
    virtual std::pair<bool, unsigned int> getRowCount(
        const QString& sheetName) = 0;

    /**
     * @brief Get data from sheet.
     * @param sheetName Sheet name.
     * @param excludedColumns Vector of excluded columns (indexes).
     * @return First value indicating success, second is vector of data rows.
     */
    std::pair<bool, QVector<QVector<QVariant>>> getData(
        const QString& sheetName, const QVector<unsigned int>& excludedColumns);

    /**
     * @brief Get limited data from sheet.
     * @param sheetName Sheet name.
     * @param excludedColumns Vector of excluded columns (indexes).
     * @param rowLimit Number of rows to load.
     * @return First value indicating success, second is vector of data rows.
     */
    virtual std::pair<bool, QVector<QVector<QVariant>>> getLimitedData(
        const QString& sheetName, const QVector<unsigned int>& excludedColumns,
        unsigned int rowLimit) = 0;

protected:
    void setError(QString functionName, QString errorContent);

    virtual void updateProgress(unsigned int currentRow, unsigned int rowCount,
                                unsigned int& lastEmittedPercent);

    bool openZip();

    bool setCurrentZipFile(const QString& zipFileName);

    bool openZipFile(QuaZipFile& zipFile);

    bool initZipFile(QuaZipFile& zipFile, const QString& zipFileName);

    QVector<QVariant> createTemplateDataRow(
        const QVector<unsigned int>& excludedColumns,
        const QVector<ColumnType>& columnTypes) const;

    QMap<unsigned int, unsigned int> createActiveColumnMapping(
        const QVector<unsigned int>& excludedColumns,
        unsigned int columnCount) const;

    bool columnsToExcludeAreValid(const QVector<unsigned int>& excludedColumns,
                                  unsigned int columnCount);

    virtual std::pair<bool, QStringList> retrieveColumnNames(
        const QString& sheetName) = 0;

    virtual std::tuple<bool, unsigned int, QVector<ColumnType>>
    retrieveRowCountAndColumnTypes(const QString& sheetName) = 0;

    bool analyzeSheet(const QString& sheetName);

    bool isSheetNameValid(const QStringList& sheetNames,
                          const QString& sheetName);

    QHash<QString, unsigned int> rowCounts_{};
    QHash<QString, unsigned int> columnCounts_{};
    QHash<QString, QVector<ColumnType>> columnTypes_{};
    QHash<QString, QStringList> columnNames_{};

    QIODevice& ioDevice_;

    /// If empty column is encountered insert defined string.
    QString emptyColName_;

    static constexpr int NOT_SET_COLUMN{-1};

private:
    std::pair<QString, QString> error_;

    QuaZip zip_;

Q_SIGNALS:
    /**
     * Triggered on change of progress percentage.
     * @param progressPercent New progress percent.
     */
    void progressPercentChanged(unsigned int progressPercent);
};

#endif  // IMPORTSPREADSHEET_H
