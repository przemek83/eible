#include "ImportOds.h"

#include <QXmlStreamReader>
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
    if (!sheetNames_ && !getSheetNames().first)
        return {false, {}};

    if (!sheetNames_->contains(sheetName))
    {
        setError(__FUNCTION__,
                 "Sheet " + sheetName +
                     " not found. Available sheets: " + sheetNames_->join(','));
        return {false, {}};
    }

    const auto it = columnCounts_.find(sheetName);
    if (it == columnCounts_.end() && !analyzeSheet(sheetName))
        return {false, {}};
    const unsigned int columnCount{*it};

    QuaZip zip(&ioDevice_);

    if (!zip.open(QuaZip::mdUnzip))
    {
        setError(__FUNCTION__,
                 "Can not open zip file " + zip.getZipName() + ".");
        return {false, {}};
    }

    QStringList columnNames{};

    if (zip.setCurrentFile(QStringLiteral("content.xml")))
    {
        // Open file in zip archive.
        QuaZipFile zipFile(&zip);
        if (!zipFile.open(QIODevice::ReadOnly | QIODevice::Text))
        {
            setError(__FUNCTION__,
                     "Can not open file " + zipFile.getFileName() + ".");
            return {false, {}};
        }

        QXmlStreamReader xmlStreamReader;
        xmlStreamReader.setDevice(&zipFile);

        // Move to first row in selected sheet.
        while (!xmlStreamReader.atEnd() &&
               xmlStreamReader.name() != "table:table" &&
               xmlStreamReader.attributes().value(
                   QLatin1String("table:name")) != sheetName)
        {
            xmlStreamReader.readNext();
        }

        while (!xmlStreamReader.atEnd() &&
               xmlStreamReader.name() != "table-row")
        {
            xmlStreamReader.readNext();
        }

        xmlStreamReader.readNext();

        // Actual column number.
        int column = NOT_SET_COLUMN;
        QXmlStreamReader::TokenType lastToken = xmlStreamReader.tokenType();

        const QString numberColumnsRepeated(
            QStringLiteral("table:number-columns-repeated"));

        // Parse first row.
        while (!xmlStreamReader.atEnd() &&
               xmlStreamReader.name() != "table-row")
        {
            // When we encounter first cell of worksheet.
            if (xmlStreamReader.name().toString() ==
                    QLatin1String("table-cell") &&
                xmlStreamReader.tokenType() == QXmlStreamReader::StartElement)
            {
                QString emptyColNumber = xmlStreamReader.attributes()
                                             .value(numberColumnsRepeated)
                                             .toString();
                if (!emptyColNumber.isEmpty())
                {
                    break;
                }

                // Add column number.
                column++;
            }

            // If we encounter start of cell content we add it to list.
            if (!xmlStreamReader.atEnd() &&
                xmlStreamReader.name().toString() == QStringLiteral("p") &&
                xmlStreamReader.tokenType() == QXmlStreamReader::StartElement)
            {
                while (xmlStreamReader.tokenType() !=
                       QXmlStreamReader::Characters)
                {
                    xmlStreamReader.readNext();
                }

                columnNames.push_back(xmlStreamReader.text().toString());
            }

            // If we encounter empty cell we add it to list.
            if (xmlStreamReader.name().toString() ==
                    QLatin1String("table-cell") &&
                xmlStreamReader.tokenType() == QXmlStreamReader::EndElement &&
                lastToken == QXmlStreamReader::StartElement)
            {
                columnNames << emptyColName_;
            }

            lastToken = xmlStreamReader.tokenType();
            xmlStreamReader.readNext();
        }
    }
    else
    {
        setError(__FUNCTION__,
                 "Can not open file " + sheetName + " in archive.");
        return {false, {}};
    }

    while (columnNames.count() < static_cast<int>(columnCount))
        columnNames << emptyColName_;

    return {true, columnNames};
}

std::pair<bool, unsigned int> ImportOds::getColumnCount(
    const QString& sheetName)
{
    return getCount(sheetName, columnCounts_);
}

std::pair<bool, unsigned int> ImportOds::getRowCount(const QString& sheetName)
{
    return getCount(sheetName, rowCounts_);
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

std::pair<bool, unsigned int> ImportOds::getCount(
    const QString& sheetName, const QHash<QString, unsigned int>& countMap)
{
    if (!sheetNames_ && !getSheetNames().first)
        return {false, {}};

    if (!sheetNames_->contains(sheetName))
    {
        setError(__FUNCTION__,
                 "Sheet " + sheetName +
                     " not found. Available sheets: " + sheetNames_->join(','));
        return {false, {}};
    }

    const auto it = countMap.find(sheetName);
    if (it != countMap.end())
        return {true, it.value()};

    if (!analyzeSheet(sheetName))
        return {false, 0};

    return {true, countMap.find(sheetName).value()};
}

bool ImportOds::analyzeSheet(const QString& sheetName)
{
    QuaZip zip(&ioDevice_);

    if (!zip.open(QuaZip::mdUnzip))
    {
        setError(__FUNCTION__,
                 "Can not open zip file " + zip.getZipName() + ".");
        return false;
    }

    QuaZipFile zipFile;
    QXmlStreamReader xmlStreamReader;

    if (!openZipAndMoveToSecondRow(zip, sheetName, zipFile, xmlStreamReader))
        return false;

    int column = NOT_SET_COLUMN;
    int rowCounter = 0;
    int repeatCount = 1;
    int maxColumn = NOT_SET_COLUMN;
    QString currentColType;
    bool rowEmpty = true;

    const QString tableTag(QStringLiteral("table"));
    const QString tableRowTag(QStringLiteral("table-row"));
    const QString tableCellTag(QStringLiteral("table-cell"));
    const QString officeValueTypeTag(QStringLiteral("office:value-type"));
    const QString columnsRepeatedTag(
        QStringLiteral("table:number-columns-repeated"));
    const QString stringTag(QStringLiteral("string"));
    const QString dateTag(QStringLiteral("date"));
    const QString floatTag(QStringLiteral("float"));
    const QString percentageTag(QStringLiteral("percentage"));
    const QString currencyTag(QStringLiteral("currency"));
    const QString timeTag(QStringLiteral("time"));

    while (!xmlStreamReader.atEnd() &&
           xmlStreamReader.name().compare(tableTag) != 0)
    {
        if (0 == xmlStreamReader.name().compare(tableRowTag) &&
            xmlStreamReader.isStartElement())
            column = NOT_SET_COLUMN;

        if (0 == xmlStreamReader.name().compare(tableCellTag) &&
            xmlStreamReader.isStartElement())
        {
            column++;

            currentColType = xmlStreamReader.attributes()
                                 .value(officeValueTypeTag)
                                 .toString();

            repeatCount = std::max(1, xmlStreamReader.attributes()
                                          .value(columnsRepeatedTag)
                                          .toString()
                                          .toInt());

            for (int i = 0; i < repeatCount; ++i)
            {
                if (0 == currentColType.compare(stringTag) ||
                    0 == currentColType.compare(dateTag) ||
                    0 == currentColType.compare(floatTag) ||
                    0 == currentColType.compare(percentageTag) ||
                    0 == currentColType.compare(currencyTag) ||
                    0 == currentColType.compare(timeTag))
                {
                    maxColumn = std::max(maxColumn, column + i);
                    rowEmpty = false;
                }
            }
            column += repeatCount - 1;
        }
        xmlStreamReader.readNext();
        if (0 == xmlStreamReader.name().compare(tableRowTag) &&
            xmlStreamReader.isEndElement())
        {
            if (!rowEmpty)
                rowCounter++;
            rowEmpty = true;
        }
    }

    rowCounts_[sheetName] = rowCounter;
    columnCounts_[sheetName] = maxColumn + 1;

    return true;
}

bool ImportOds::openZipAndMoveToSecondRow(QuaZip& zip, const QString& sheetName,
                                          QuaZipFile& zipFile,
                                          QXmlStreamReader& xmlStreamReader)
{
    // Open file in zip archive.
    if (!zip.setCurrentFile(QStringLiteral("content.xml")))
    {
        setError(__FUNCTION__,
                 "Can not open file " + sheetName + " in archive.");
        return false;
    }
    zipFile.setZip(&zip);
    if (!zipFile.open(QIODevice::ReadOnly | QIODevice::Text))
    {
        setError(__FUNCTION__,
                 "Can not open file " + zipFile.getFileName() + ".");
        return false;
    }

    xmlStreamReader.setDevice(&zipFile);

    // Move to first row in selected sheet.
    while (!xmlStreamReader.atEnd() &&
           xmlStreamReader.name() != "table:table" &&
           xmlStreamReader.attributes().value(QLatin1String("table:name")) !=
               sheetName)
        xmlStreamReader.readNext();

    bool secondRow = false;
    while (!xmlStreamReader.atEnd() &&
           !(xmlStreamReader.name() == "table" &&
             xmlStreamReader.tokenType() == QXmlStreamReader::EndElement))
    {
        if (xmlStreamReader.name() == "table-row" &&
            xmlStreamReader.tokenType() == QXmlStreamReader::StartElement)
        {
            if (secondRow)
                break;
            secondRow = true;
        }
        xmlStreamReader.readNext();
    }

    return true;
}
