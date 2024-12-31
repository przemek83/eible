#include <eible/ImportOds.h>

#include <cmath>
#include <utility>

#include <quazip/quazip.h>
#include <quazip/quazipfile.h>
#include <QDebug>
#include <QVariant>
#include <QVector>
#include <QXmlStreamReader>
#include <QtXml/QDomDocument>

ImportOds::ImportOds(QIODevice& ioDevice) : ImportSpreadsheet(ioDevice) {}

std::pair<bool, QStringList> ImportOds::getSheetNames()
{
    if (sheetNames_)
        return {true, *sheetNames_};

    QuaZipFile zipFile;
    if (!initZipFile(zipFile, QStringLiteral("settings.xml")))
        return {false, {}};

    QDomDocument xmlDocument;
    if (!xmlDocument.setContent(zipFile.readAll()))
    {
        setError(QStringLiteral("Xml file is damaged."));
        return {false, {}};
    }

    const QString configMapNamed{
        QStringLiteral("config:config-item-map-named")};
    const QString configName{QStringLiteral("config:name")};
    const QString configMapEntry{
        QStringLiteral("config:config-item-map-entry")};
    const QString tables{QStringLiteral("Tables")};

    const QDomElement root{xmlDocument.documentElement()};
    const int elementsCount{root.elementsByTagName(configMapNamed).size()};
    QStringList sheetNames{};
    for (int i{0}; i < elementsCount; ++i)
    {
        const QDomElement currentElement{
            root.elementsByTagName(configMapNamed).at(i).toElement()};
        if (currentElement.hasAttribute(configName) &&
            currentElement.attribute(configName) == tables)
        {
            const int innerElementsCount{
                currentElement.elementsByTagName(configMapEntry).size()};
            for (int j{0}; j < innerElementsCount; ++j)
            {
                const QDomElement element{
                    currentElement.elementsByTagName(configMapEntry)
                        .at(j)
                        .toElement()};
                sheetNames.push_back(element.attribute(configName));
            }
        }
    }

    sheetNames_ = std::move(sheetNames);
    return {true, *sheetNames_};
}

std::pair<bool, QVector<ColumnType>> ImportOds::getColumnTypes(
    const QString& sheetName)
{
    if (const auto it{columnTypes_.constFind(sheetName)};
        it != columnTypes_.constEnd())
        return {true, *it};

    if (!sheetNames_ && !getSheetNames().first)
        return {false, {}};

    if (sheetNames_.has_value() && !isSheetNameValid(*sheetNames_, sheetName))
        return {false, {}};

    QuaZipFile zipFile;
    if (!initZipFile(zipFile, QStringLiteral("content.xml")))
        return {false, {}};

    QXmlStreamReader xmlStreamReader;
    if (!moveToSecondRow(sheetName, zipFile, xmlStreamReader))
        return {false, {}};

    if (!analyzeSheet(sheetName))
        return {false, {}};

    return {true, columnTypes_.value(sheetName)};
}

std::pair<bool, QStringList> ImportOds::getColumnNames(const QString& sheetName)
{
    if (!sheetNames_ && !getSheetNames().first)
        return {false, {}};

    if (sheetNames_.has_value() && !isSheetNameValid(*sheetNames_, sheetName))
        return {false, {}};

    const auto it{columnNames_.constFind(sheetName)};
    if (it == columnNames_.constEnd() && !analyzeSheet(sheetName))
        return {false, {}};

    return {true, columnNames_.value(sheetName)};
}

std::pair<bool, QVector<QVector<QVariant>>> ImportOds::getLimitedData(
    const QString& sheetName, const QVector<unsigned int>& excludedColumns,
    unsigned int rowLimit)
{
    const auto [success, columnTypes]{getColumnTypes(sheetName)};
    if (!success)
        return {false, {}};

    const unsigned int columnCount{columnCounts_.value(sheetName)};
    if (!columnsToExcludeAreValid(excludedColumns, columnCount))
        return {false, {}};

    QuaZipFile zipFile;
    if (!initZipFile(zipFile, QStringLiteral("content.xml")))
        return {false, {}};

    QXmlStreamReader xmlStreamReader;
    if (!moveToSecondRow(sheetName, zipFile, xmlStreamReader))
        return {false, {}};

    const QVector<QVariant> templateDataRow(
        createTemplateDataRow(excludedColumns, columnTypes));

    QMap<unsigned int, unsigned int> activeColumnsMapping{
        createActiveColumnMapping(excludedColumns, columnCount)};

    QVector<QVector<QVariant>> dataContainer(
        static_cast<int>(std::min(getRowCount(sheetName).second, rowLimit)));

    QVector<QVariant> currentDataRow(templateDataRow);

    int column{NOT_SET_COLUMN};

    unsigned int rowCounter{0};
    bool rowEmpty{true};
    unsigned int lastEmittedPercent{0};
    while (!xmlStreamReader.atEnd() &&
           0 != xmlStreamReader.name().compare(TABLE_TAG) &&
           rowCounter < rowLimit)
    {
        if (isRowStart(xmlStreamReader))
        {
            column = NOT_SET_COLUMN;
            if (!rowEmpty)
            {
                dataContainer[static_cast<int>(rowCounter)] = currentDataRow;
                currentDataRow = templateDataRow;
                ++rowCounter;
                updateProgress(rowCounter, rowLimit, lastEmittedPercent);
            }
            rowEmpty = true;
        }

        if (isCellStart(xmlStreamReader))
        {
            ++column;
            const unsigned int repeats{
                getColumnRepeatCount(xmlStreamReader.attributes())};
            if (!isOfficeValueTagEmpty(xmlStreamReader))
            {
                rowEmpty = false;
                const QVariant value{retrieveValueFromField(
                    xmlStreamReader, columnTypes.at(column))};
                for (unsigned int i{0}; i < repeats; ++i)
                {
                    const unsigned int currentColumn{
                        static_cast<unsigned int>(column) + i};
                    if (excludedColumns.contains(currentColumn))
                        continue;
                    const unsigned int mappedColumnIndex{
                        activeColumnsMapping[currentColumn]};
                    currentDataRow[static_cast<int>(mappedColumnIndex)] = value;
                }
            }
            column += static_cast<int>(repeats - 1);
        }
        xmlStreamReader.readNextStartElement();
    }

    if (currentDataRow != templateDataRow)
        dataContainer[dataContainer.size() - 1] = currentDataRow;

    return {true, dataContainer};
}

std::pair<bool, unsigned int> ImportOds::getCount(
    const QString& sheetName, const QHash<QString, unsigned int>& countMap)
{
    if (!sheetNames_ && !getSheetNames().first)
        return {false, {}};

    if (sheetNames_.has_value() && !isSheetNameValid(*sheetNames_, sheetName))
        return {false, {}};

    if (const auto it{countMap.find(sheetName)}; it != countMap.end())
        return {true, it.value()};

    if (!analyzeSheet(sheetName))
        return {false, 0};

    return {true, countMap.find(sheetName).value()};
}

std::pair<bool, QStringList> ImportOds::retrieveColumnNames(
    const QString& sheetName)
{
    QuaZipFile zipFile;
    if (!initZipFile(zipFile, QStringLiteral("content.xml")))
        return {false, {}};

    QXmlStreamReader xmlStreamReader;
    xmlStreamReader.setDevice(&zipFile);
    skipToSheet(xmlStreamReader, sheetName);

    while ((!xmlStreamReader.atEnd()) &&
           (xmlStreamReader.name() != TABLE_ROW_TAG))
        xmlStreamReader.readNext();
    xmlStreamReader.readNext();

    QXmlStreamReader::TokenType lastToken{xmlStreamReader.tokenType()};
    QStringList columnNames;

    while (!xmlStreamReader.atEnd() &&
           (xmlStreamReader.name() != TABLE_ROW_TAG))
    {
        if (isCellStart(xmlStreamReader) &&
            (getColumnRepeatCount(xmlStreamReader.attributes()) > 1))
            break;

        if ((xmlStreamReader.name().toString() == P_TAG) &&
            xmlStreamReader.isStartElement())
        {
            while (xmlStreamReader.tokenType() != QXmlStreamReader::Characters)
                xmlStreamReader.readNext();
            columnNames.push_back(xmlStreamReader.text().toString());
        }

        if (isCellEnd(xmlStreamReader) &&
            (lastToken == QXmlStreamReader::StartElement))
            columnNames << emptyColName_;

        lastToken = xmlStreamReader.tokenType();
        xmlStreamReader.readNext();
    }
    return {true, columnNames};
}

std::tuple<bool, unsigned int, QVector<ColumnType>>
ImportOds::retrieveRowCountAndColumnTypes(const QString& sheetName)
{
    QuaZipFile zipFile;
    if (!initZipFile(zipFile, QStringLiteral("content.xml")))
        return {false, {}, {}};

    QXmlStreamReader xmlStreamReader;
    if (!moveToSecondRow(sheetName, zipFile, xmlStreamReader))
        return {false, {}, {}};

    QVector<ColumnType> columnTypes;
    int column{NOT_SET_COLUMN};
    int rowCounter{0};
    bool rowEmpty{true};

    while (!xmlStreamReader.atEnd() &&
           xmlStreamReader.name().compare(TABLE_TAG) != 0)
    {
        if (isRowStart(xmlStreamReader))
            column = NOT_SET_COLUMN;

        if (isCellStart(xmlStreamReader))
        {
            ++column;
            const QXmlStreamAttributes attributes{xmlStreamReader.attributes()};
            const QString xmlColTypeValue{
                attributes.value(OFFICE_VALUE_TYPE_TAG).toString()};
            const unsigned int repeats{getColumnRepeatCount(attributes)};
            if (isRecognizedColumnType(attributes))
            {
                rowEmpty = false;
                for (int i{static_cast<int>(columnTypes.size())};
                     i < (column + static_cast<int>(repeats)); ++i)
                    columnTypes.push_back(ColumnType::UNKNOWN);

                for (unsigned int i{0}; i < repeats; ++i)
                {
                    const int currentColumn{column + static_cast<int>(i)};
                    columnTypes[currentColumn] = recognizeColumnType(
                        columnTypes.at(currentColumn), xmlColTypeValue);
                }
            }
            column += static_cast<int>(repeats - 1U);
        }
        xmlStreamReader.readNext();
        if (isRowEnd(xmlStreamReader))
        {
            if (!rowEmpty)
                ++rowCounter;
            rowEmpty = true;
        }
    }

    std::replace_if(
        columnTypes.begin(), columnTypes.end(), [](ColumnType columnType)
        { return columnType == ColumnType::UNKNOWN; }, ColumnType::STRING);

    return {true, rowCounter, columnTypes};
}

bool ImportOds::moveToSecondRow(const QString& sheetName, QuaZipFile& zipFile,
                                QXmlStreamReader& xmlStreamReader) const
{
    xmlStreamReader.setDevice(&zipFile);
    skipToSheet(xmlStreamReader, sheetName);

    bool secondRow{false};
    while (!xmlStreamReader.atEnd() &&
           !(xmlStreamReader.qualifiedName() == TABLE_QUALIFIED_NAME &&
             xmlStreamReader.isEndElement()))
    {
        if (isRowStart(xmlStreamReader))
        {
            if (secondRow)
                break;
            secondRow = true;
        }
        xmlStreamReader.readNext();
    }
    return true;
}

void ImportOds::skipToSheet(QXmlStreamReader& xmlStreamReader,
                            const QString& sheetName) const
{
    while (!xmlStreamReader.atEnd())
    {
        if (xmlStreamReader.qualifiedName() == TABLE_QUALIFIED_NAME)
        {
            if (xmlStreamReader.attributes().value(TABLE_NAME_TAG) != sheetName)
                xmlStreamReader.skipCurrentElement();
            else
                break;
        }
        xmlStreamReader.readNextStartElement();
    }
}

bool ImportOds::isRecognizedColumnType(
    const QXmlStreamAttributes& attributes) const
{
    const QString columnType{
        attributes.value(OFFICE_VALUE_TYPE_TAG).toString()};
    return RECOCNIZED_COLUMN_TYPES.contains(columnType);
}

unsigned int ImportOds::getColumnRepeatCount(
    const QXmlStreamAttributes& attributes) const
{
    return std::max(1U,
                    attributes.value(COLUMNS_REPEATED_TAG).toString().toUInt());
}

bool ImportOds::isRowStart(const QXmlStreamReader& xmlStreamReader) const
{
    return (0 == xmlStreamReader.name().compare(TABLE_ROW_TAG)) &&
           xmlStreamReader.isStartElement();
}

bool ImportOds::isRowEnd(const QXmlStreamReader& xmlStreamReader) const
{
    return (0 == xmlStreamReader.name().compare(TABLE_ROW_TAG)) &&
           xmlStreamReader.isEndElement();
}

bool ImportOds::isCellStart(const QXmlStreamReader& xmlStreamReader) const
{
    return (0 == xmlStreamReader.name().compare(TABLE_CELL_TAG)) &&
           xmlStreamReader.isStartElement();
}

bool ImportOds::isCellEnd(const QXmlStreamReader& xmlStreamReader) const
{
    return (xmlStreamReader.name().toString() == TABLE_CELL_TAG) &&
           xmlStreamReader.isEndElement();
}

ColumnType ImportOds::recognizeColumnType(ColumnType currentType,
                                          const QString& xmlColTypeValue) const
{
    if (0 == xmlColTypeValue.compare(STRING_TAG))
    {
        if (currentType == ColumnType::UNKNOWN)
            return ColumnType::STRING;
        if (currentType != ColumnType::STRING)
            return ColumnType::STRING;
    }

    if (0 == xmlColTypeValue.compare(DATE_TAG))
    {
        if (currentType == ColumnType::UNKNOWN)
            return ColumnType::DATE;
        if (currentType != ColumnType::DATE)
            return ColumnType::STRING;
    }

    if ((0 == xmlColTypeValue.compare(FLOAT_TAG)) ||
        (0 == xmlColTypeValue.compare(PERCENTAGE_TAG)) ||
        (0 == xmlColTypeValue.compare(CURRENCY_TAG)) ||
        (0 == xmlColTypeValue.compare(TIME_TAG)))
    {
        if (currentType == ColumnType::UNKNOWN)
            return ColumnType::NUMBER;
        if (currentType != ColumnType::NUMBER)
            return ColumnType::STRING;
    }

    return currentType;
}

QVariant ImportOds::retrieveValueFromField(QXmlStreamReader& xmlStreamReader,
                                           ColumnType columnType) const
{
    const QString xmlColTypeValue{
        xmlStreamReader.attributes().value(OFFICE_VALUE_TYPE_TAG).toString()};
    QVariant value{};
    const QString emptyString(QLatin1String(""));

    switch (columnType)
    {
        case ColumnType::STRING:
        {
            const QString currentDateValue{xmlStreamReader.attributes()
                                               .value(OFFICE_DATE_VALUE_TAG)
                                               .toString()};

            while (!xmlStreamReader.atEnd() &&
                   0 != xmlStreamReader.name().compare(P_TAG))
                xmlStreamReader.readNext();

            while (xmlStreamReader.tokenType() !=
                       QXmlStreamReader::Characters &&
                   0 != xmlStreamReader.name().compare(TABLE_CELL_TAG))
                xmlStreamReader.readNext();
            if (xmlColTypeValue.compare(DATE_TAG) == 0)
                value = QVariant(currentDateValue);
            else
            {
                const QStringView stringView{xmlStreamReader.text()};
                if (stringView.isNull())
                    value = QVariant(emptyString);
                else
                    value = stringView.toString();
            }
            break;
        }

        case ColumnType::DATE:
        {
            static const int odsStringDateLength{10};
            QString dateValue{xmlStreamReader.attributes()
                                  .value(OFFICE_DATE_VALUE_TAG)
                                  .toString()};
            dateValue.chop(dateValue.size() - odsStringDateLength);
            value = QDate::fromString(dateValue, DATE_FORMAT);

            break;
        }

        case ColumnType::NUMBER:
        {
            value = QVariant(xmlStreamReader.attributes()
                                 .value(OFFICE_VALUE_TAG)
                                 .toDouble());
            break;
        }

        case ColumnType::UNKNOWN:
        {
            Q_ASSERT(false);
            break;
        }
    }
    return value;
}

bool ImportOds::isOfficeValueTagEmpty(
    const QXmlStreamReader& xmlStreamReader) const
{
    const QString xmlColTypeValue{
        xmlStreamReader.attributes().value(OFFICE_VALUE_TYPE_TAG).toString()};
    return xmlColTypeValue.isEmpty();
}
