#ifndef IMPORTSPREADSHEET_H
#define IMPORTSPREADSHEET_H

#include <QObject>

#include <quazip5/quazip.h>

#include "ColumnType.h"
#include "eible_global.h"

class QIODevice;
class QuaZipFile;

class EIBLE_EXPORT ImportSpreadsheet : public QObject
{
    Q_OBJECT
public:
    explicit ImportSpreadsheet(QIODevice& ioDevice);

    virtual std::pair<bool, QStringList> getSheetNames() = 0;

    virtual std::pair<bool, QStringList> getColumnNames(
        const QString& sheetName) = 0;

    virtual std::pair<bool, QVector<ColumnType>> getColumnTypes(
        const QString& sheetName) = 0;

    std::pair<QString, QString> getError() const;

    void setNameForEmptyColumn(const QString& name);

    virtual std::pair<bool, unsigned int> getColumnCount(
        const QString& sheetName) = 0;

    virtual std::pair<bool, unsigned int> getRowCount(
        const QString& sheetName) = 0;

    std::pair<bool, QVector<QVector<QVariant>>> getData(
        const QString& sheetName, const QVector<unsigned int>& excludedColumns);

    virtual std::pair<bool, QVector<QVector<QVariant>>> getLimitedData(
        const QString& sheetName, const QVector<unsigned int>& excludedColumns,
        unsigned int rowLimit) = 0;

protected:
    void setError(QString functionName, QString errorContent);

    virtual void updateProgress(unsigned int currentRow, unsigned int rowCount,
                                unsigned int& lastEmittedPercent);

    bool openZipFile(QuaZipFile& zipFile, const QString& zipFileName);

    QIODevice& ioDevice_;

    /// If empty column is encountered insert defined string.
    QString emptyColName_;

    static constexpr int NOT_SET_COLUMN{-1};

private:
    std::pair<QString, QString> error_;

    QuaZip zip_;

Q_SIGNALS:
    void progressPercentChanged(unsigned int progressPercent);
};

#endif  // IMPORTSPREADSHEET_H
