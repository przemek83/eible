#ifndef IMPORTSPREADSHEET_H
#define IMPORTSPREADSHEET_H

#include <QObject>

#include "ColumnType.h"
#include "eible_global.h"

class QIODevice;

class EIBLE_EXPORT ImportSpreadsheet : public QObject
{
    Q_OBJECT
public:
    explicit ImportSpreadsheet(QIODevice& ioDevice);

    virtual std::pair<bool, QStringList> getSheetNames() = 0;

    virtual std::pair<bool, QStringList> getColumnList(
        const QString& sheetName) = 0;

    virtual std::pair<bool, QVector<ColumnType>> getColumnTypes(
        const QString& sheetName, int columnsCount) = 0;

    std::pair<QString, QString> getError() const;

    void setNameForEmptyColumn(const QString& name);

protected:
    void setError(QString functionName, QString errorContent);

    QIODevice& ioDevice_;

    /// If empty column is encountered insert defined string.
    QString emptyColName_;

private:
    std::pair<QString, QString> error_;

Q_SIGNALS:
    void updateProgress(int progress);
};

#endif  // IMPORTSPREADSHEET_H
