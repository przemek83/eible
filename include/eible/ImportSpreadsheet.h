#pragma once

#include <quazip/quazip.h>
#include <QObject>
#include <QVector>

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
    QString getLastError() const;

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
    std::pair<bool, int> getColumnCount(const QString& sheetName);

    /**
     * @brief Get number of rows in given sheet.
     * @param sheetName Sheet name.
     * @return First value indicating success, second is row count.
     */
    std::pair<bool, int> getRowCount(const QString& sheetName);

    /**
     * @brief Get data from sheet.
     * @param sheetName Sheet name.
     * @param excludedColumns Vector of excluded columns (indexes).
     * @return First value indicating success, second is vector of data rows.
     */
    std::pair<bool, QVector<QVector<QVariant>>> getData(
        const QString& sheetName, const QVector<int>& excludedColumns);

    /**
     * @brief Get limited data from sheet.
     * @param sheetName Sheet name.
     * @param excludedColumns Vector of excluded columns (indexes).
     * @param rowLimit Number of rows to load.
     * @return First value indicating success, second is vector of data rows.
     */
    virtual std::pair<bool, QVector<QVector<QVariant>>> getLimitedData(
        const QString& sheetName, const QVector<int>& excludedColumns,
        int rowLimit) = 0;

protected:
    void setError(const QString& errorContent);

    virtual void updateProgress(int currentRow, int rowCount,
                                int& lastEmittedPercent);

    bool openZip();

    bool setCurrentZipFile(const QString& zipFileName);

    bool openZipFile(QuaZipFile& quaZipFile);

    bool initZipFile(QuaZipFile& quaZipFile, const QString& zipFileName);

    static QVector<QVariant> createTemplateDataRow(
        const QVector<int>& excludedColumns,
        const QVector<ColumnType>& columnTypes);

    static QMap<int, int> createActiveColumnMapping(
        const QVector<int>& excludedColumns, int columnCount);

    bool columnsToExcludeAreValid(const QVector<int>& excludedColumns,
                                  int columnCount);

    virtual std::pair<bool, QStringList> retrieveColumnNames(
        const QString& sheetName) = 0;

    virtual std::tuple<bool, int, QVector<ColumnType>>
    retrieveRowCountAndColumnTypes(const QString& sheetName) = 0;

    bool analyzeSheet(const QString& sheetName);

    bool isSheetNameValid(const QStringList& sheetNames,
                          const QString& sheetName);

    virtual std::pair<bool, int> getCount(
        const QString& sheetName, const QHash<QString, int>& countMap) = 0;

    QHash<QString, int> rowCounts_{};
    QHash<QString, int> columnCounts_{};
    QHash<QString, QVector<ColumnType>> columnTypes_{};
    QHash<QString, QStringList> columnNames_{};

    QIODevice& ioDevice_;

    /// If empty column is encountered than insert defined string.
    QString emptyColName_;

    static constexpr int NOT_SET_COLUMN{-1};

private:
    QString lastError_;

    QuaZip zip_;

Q_SIGNALS:
    /**
     * Triggered on change of progress percentage.
     * @param progressPercent New progress percent.
     */
    void progressPercentChanged(int progressPercent);
};
