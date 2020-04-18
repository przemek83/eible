#ifndef IMPORTSPREADSHEET_H
#define IMPORTSPREADSHEET_H

#include <QObject>

#include "eible_global.h"

class QIODevice;

class EIBLE_EXPORT ImportSpreadsheet : public QObject
{
    Q_OBJECT
public:
    explicit ImportSpreadsheet(QIODevice& ioDevice);

    virtual std::pair<bool, QMap<QString, QString>> getSheetList() = 0;

    virtual std::pair<bool, QStringList> getColumnList(
        const QString& sheetName, QHash<QString, int> sharedStrings) = 0;

    std::pair<QString, QString> getError() const;

protected:
    void setError(QString functionName, QString errorContent);

    QIODevice& ioDevice_;

    /// If empty column is encountered insert defined string.
    const QString emptyColName_;

private:
    std::pair<QString, QString> error_;

Q_SIGNALS:
    void updateProgress(int progress);
};

#endif  // IMPORTSPREADSHEET_H
