#ifndef IMPORTSPREADSHEET_H
#define IMPORTSPREADSHEET_H

#include <QObject>

class QIODevice;

class ImportSpreadsheet : public QObject
{
    Q_OBJECT
public:
    explicit ImportSpreadsheet(QIODevice& ioDevice);

    virtual std::pair<bool, QMap<QString, QString>> getSheetList() = 0;

    std::pair<QString, QString> getError() const;

protected:
    void setError(QString functionName, QString errorContent);

    QIODevice& ioDevice_;

private:
    std::pair<QString, QString> error_;

Q_SIGNALS:
    void updateProgress(int progress);
};

#endif  // IMPORTSPREADSHEET_H
