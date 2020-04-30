#include "ImportOds.h"

#include <utility>

#include <quazip5/quazip.h>
#include <quazip5/quazipfile.h>
#include <QVariant>
#include <QVector>
#include <QtXml/QDomDocument>

ImportOds::ImportOds(QIODevice& ioDevice) : ImportSpreadsheet(ioDevice) {}

std::pair<bool, QStringList> ImportOds::getSheetNames()
{
    if (sheetNames_)
        return {true, *sheetNames_};

    QuaZip zip(&ioDevice_);

    if (!zip.open(QuaZip::mdUnzip))
    {
        setError(__FUNCTION__,
                 "Can not open zip file " + zip.getZipName() + ".");
        return {false, {}};
    }

    QStringList sheetNames{};
    const QString innerFileToOpen{QStringLiteral("settings.xml")};
    if (zip.setCurrentFile(innerFileToOpen))
    {
        // Open file in zip archive.
        QuaZipFile zipFile(&zip);
        if (!zipFile.open(QIODevice::ReadOnly | QIODevice::Text))
        {
            setError(__FUNCTION__,
                     "Can not open file " + zipFile.getFileName() + ".");
            return {false, {}};
        }

        // Create, set content and read DOM.
        QDomDocument xmlDocument(__FUNCTION__);
        if (!xmlDocument.setContent(zipFile.readAll()))
        {
            setError(__FUNCTION__, "Xml file is damaged.");
            return {false, {}};
        }
        zipFile.close();

        QDomElement root = xmlDocument.documentElement();

        const QString configMapNamed(
            QStringLiteral("config:config-item-map-named"));
        const QString configName(QStringLiteral("config:name"));
        const QString configMapEntry(
            QStringLiteral("config:config-item-map-entry"));
        const QString tables(QStringLiteral("Tables"));

        int elementsCount = root.elementsByTagName(configMapNamed).size();
        for (int i = 0; i < elementsCount; i++)
        {
            QDomElement currentElement =
                root.elementsByTagName(configMapNamed).at(i).toElement();
            if (currentElement.hasAttribute(configName) &&
                currentElement.attribute(configName) == tables)
            {
                int innerElementsCount =
                    currentElement.elementsByTagName(configMapEntry).size();
                for (int j = 0; j < innerElementsCount; j++)
                {
                    QDomElement element =
                        currentElement.elementsByTagName(configMapEntry)
                            .at(j)
                            .toElement();
                    sheetNames.push_back(element.attribute(configName));
                }
            }
        }
    }
    else
    {
        setError(__FUNCTION__,
                 "Can not open file " + innerFileToOpen + " in archive.");
        return {false, {}};
    }

    sheetNames_ = std::move(sheetNames);
    return {true, *sheetNames_};
}

std::pair<bool, QVector<ColumnType> > ImportOds::getColumnTypes(
    const QString& sheetName)
{
    return {false, {}};
}

std::pair<bool, QStringList> ImportOds::getColumnNames(const QString& sheetName)
{
    return {false, {}};
}

std::pair<bool, unsigned int> ImportOds::getColumnCount(
    const QString& sheetName)
{
    return {false, {}};
}

std::pair<bool, unsigned int> ImportOds::getRowCount(const QString& sheetName)
{
    return {false, {}};
}

std::pair<bool, QVector<QVector<QVariant> > > ImportOds::getData(
    const QString& sheetName, const QVector<unsigned int>& excludedColumns)
{
    return {false, {}};
}

std::pair<bool, QVector<QVector<QVariant> > > ImportOds::getLimitedData(
    const QString& sheetName, const QVector<unsigned int>& excludedColumns,
    unsigned int rowLimit)
{
    return {false, {}};
}
