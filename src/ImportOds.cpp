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
        return {true, sheetNames_.value()};

    auto [success, sheetNames]{getSheetNamesFromZipFile()};
    if (!success)
        return {false, {}};

    sheetNames_ = std::move(sheetNames);
    return {true, sheetNames_.value()};
}

std::pair<bool, QVector<ColumnType>> ImportOds::getColumnTypes(
    const QString& sheetName)
{
    if (const auto it{columnTypes_.constFind(sheetName)};
        it != columnTypes_.constEnd())
        return {true, *it};

    if (isSheetAvailable(sheetName) && initializeColumnTypes(sheetName))
        return {true, columnTypes_.value(sheetName)};

    return {false, {}};
}

std::pair<bool, QStringList> ImportOds::getColumnNames(const QString& sheetName)
{
    if (const auto it{columnNames_.constFind(sheetName)};
        it != columnNames_.constEnd())
        return {true, *it};

    if (isSheetAvailable(sheetName) && analyzeSheet(sheetName))
        return {true, columnNames_.value(sheetName)};

    return {false, {}};
}

std::pair<bool, QVector<QVector<QVariant>>> ImportOds::getLimitedData(
    const QString& sheetName, const QVector<int>& excludedColumns, int rowLimit)
{
    const auto [success, columnTypes]{getColumnTypes(sheetName)};
    if (!success)
        return {false, {}};

    const int columnCount{columnCounts_.value(sheetName)};
    if (!columnsToExcludeAreValid(excludedColumns, columnCount))
        return {false, {}};

    QuaZipFile quaZipFile;
    if (!initZipFile(quaZipFile, QStringLiteral("content.xml")))
        return {false, {}};

    QXmlStreamReader reader;
    moveToSecondRow(sheetName, quaZipFile, reader);

    const QVector<QVariant> templateDataRow(
        createTemplateDataRow(excludedColumns, columnTypes));

    QMap<int, int> activeColumnsMapping{
        createActiveColumnMapping(excludedColumns, columnCount)};

    QVector<QVector<QVariant>> dataContainer(
        std::min(getRowCount(sheetName).second, rowLimit));

    QVector<QVariant> currentDataRow(templateDataRow);

    int column{NOT_SET_COLUMN};

    int rowCounter{0};
    bool rowEmpty{true};
    int lastEmittedPercent{0};
    while ((!reader.atEnd()) && (reader.name() != TABLE_TAG) &&
           (rowCounter < rowLimit))
    {
        if (isRowStart(reader))
        {
            column = NOT_SET_COLUMN;
            if (!rowEmpty)
            {
                dataContainer[rowCounter] = currentDataRow;
                currentDataRow = templateDataRow;
                ++rowCounter;
                updateProgress(rowCounter, rowLimit, lastEmittedPercent);
            }
            rowEmpty = true;
        }

        if (isCellStart(reader))
        {
            ++column;
            const int repeats{getColumnRepeatCount(reader.attributes())};
            if (!isOfficeValueTagEmpty(reader))
            {
                rowEmpty = false;
                const QVariant value{
                    retrieveValueFromField(reader, columnTypes.at(column))};
                for (int i{0}; i < repeats; ++i)
                {
                    const int currentColumn{column + i};
                    if (!excludedColumns.contains(currentColumn))
                    {
                        const int mappedColumnIndex{
                            activeColumnsMapping[currentColumn]};
                        currentDataRow[mappedColumnIndex] = value;
                    }
                }
            }
            column += repeats - 1;
        }
        reader.readNextStartElement();
    }

    if (currentDataRow != templateDataRow)
        dataContainer[dataContainer.size() - 1] = currentDataRow;

    return {true, dataContainer};
}

std::pair<bool, QStringList> ImportOds::getSheetNamesFromZipFile()
{
    QuaZipFile quaZipFile;
    if (!initZipFile(quaZipFile, QStringLiteral("settings.xml")))
        return {false, {}};

    QDomDocument xmlDocument;
    if (!xmlDocument.setContent(quaZipFile.readAll()))
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
            (currentElement.attribute(configName) == tables))
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

    return {true, sheetNames};
}

std::pair<bool, int> ImportOds::getCount(const QString& sheetName,
                                         const QHash<QString, int>& countMap)
{
    if (const auto it{countMap.constFind(sheetName)}; it != countMap.constEnd())
        return {true, *it};

    if (isSheetAvailable(sheetName) && analyzeSheet(sheetName))
        return {true, countMap.value(sheetName)};

    return {false, {}};
}

std::pair<bool, QStringList> ImportOds::retrieveColumnNames(
    const QString& sheetName)
{
    QuaZipFile quaZipFile;
    if (!initZipFile(quaZipFile, QStringLiteral("content.xml")))
        return {false, {}};

    QXmlStreamReader reader;
    reader.setDevice(&quaZipFile);
    skipToSheet(reader, sheetName);

    while ((!reader.atEnd()) && (reader.name() != TABLE_ROW_TAG))
        reader.readNext();
    reader.readNext();

    QXmlStreamReader::TokenType lastToken{reader.tokenType()};
    QStringList columnNames;

    while ((!reader.atEnd()) && (reader.name() != TABLE_ROW_TAG))
    {
        if (isCellStart(reader) &&
            (getColumnRepeatCount(reader.attributes()) > 1))
            return {true, columnNames};

        if ((reader.name().toString() == P_TAG) && reader.isStartElement())
        {
            while (reader.tokenType() != QXmlStreamReader::Characters)
                reader.readNext();
            columnNames.push_back(reader.text().toString());
        }

        if (isCellEnd(reader) && (lastToken == QXmlStreamReader::StartElement))
            columnNames << emptyColName_;

        lastToken = reader.tokenType();
        reader.readNext();
    }
    return {true, columnNames};
}

void ImportOds::updateColumnTypes(QVector<ColumnType>& columnTypes, int column,
                                  const QString& xmlColTypeValue,
                                  int repeats) const
{
    for (qsizetype i{columnTypes.size()}; i < (column + repeats); ++i)
        columnTypes.push_back(ColumnType::UNKNOWN);

    for (int i{0}; i < repeats; ++i)
    {
        const int currentColumn{column + i};
        columnTypes[currentColumn] =
            recognizeColumnType(columnTypes.at(currentColumn), xmlColTypeValue);
    }
}

std::tuple<bool, int, QVector<ColumnType>>
ImportOds::retrieveRowCountAndColumnTypes(const QString& sheetName)
{
    QuaZipFile quaZipFile;
    if (!initZipFile(quaZipFile, QStringLiteral("content.xml")))
        return {false, {}, {}};

    QXmlStreamReader reader;
    moveToSecondRow(sheetName, quaZipFile, reader);

    QVector<ColumnType> columnTypes;
    int column{NOT_SET_COLUMN};
    int rowCounter{0};
    bool rowEmpty{true};

    while ((!reader.atEnd()) && (reader.name() != TABLE_TAG))
    {
        if (isRowStart(reader))
            column = NOT_SET_COLUMN;

        if (isCellStart(reader))
        {
            ++column;
            const QXmlStreamAttributes attributes{reader.attributes()};
            const QString xmlColTypeValue{
                attributes.value(OFFICE_VALUE_TYPE_TAG).toString()};
            const int repeats{getColumnRepeatCount(attributes)};
            if (isRecognizedColumnType(attributes))
            {
                rowEmpty = false;
                updateColumnTypes(columnTypes, column, xmlColTypeValue,
                                  repeats);
            }
            column += repeats - 1;
        }
        reader.readNext();
        if (isRowEnd(reader))
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

void ImportOds::moveToSecondRow(const QString& sheetName,
                                QuaZipFile& quaZipFile,
                                QXmlStreamReader& reader) const
{
    reader.setDevice(&quaZipFile);
    skipToSheet(reader, sheetName);

    bool secondRow{false};
    while ((!reader.atEnd()) &&
           (!((reader.qualifiedName() == TABLE_QUALIFIED_NAME) &&
              reader.isEndElement())))
    {
        if (isRowStart(reader))
        {
            if (secondRow)
                return;

            secondRow = true;
        }
        reader.readNext();
    }
}

void ImportOds::skipToSheet(QXmlStreamReader& reader,
                            const QString& sheetName) const
{
    while (!reader.atEnd())
    {
        if (reader.qualifiedName() == TABLE_QUALIFIED_NAME)
        {
            if (reader.attributes().value(TABLE_NAME_TAG) != sheetName)
                reader.skipCurrentElement();
            else
                return;
        }
        reader.readNextStartElement();
    }
}

bool ImportOds::isRecognizedColumnType(
    const QXmlStreamAttributes& attributes) const
{
    const QString columnType{
        attributes.value(OFFICE_VALUE_TYPE_TAG).toString()};
    return RECOCNIZED_COLUMN_TYPES.contains(columnType);
}

int ImportOds::getColumnRepeatCount(
    const QXmlStreamAttributes& attributes) const
{
    const int repeats{
        attributes.value(COLUMNS_REPEATED_TAG).toString().toInt()};
    return std::max(1, repeats);
}

bool ImportOds::isRowStart(const QXmlStreamReader& reader) const
{
    return (reader.name() == TABLE_ROW_TAG) && reader.isStartElement();
}

bool ImportOds::isRowEnd(const QXmlStreamReader& reader) const
{
    return (reader.name() == TABLE_ROW_TAG) && reader.isEndElement();
}

bool ImportOds::isCellStart(const QXmlStreamReader& reader) const
{
    return (reader.name() == TABLE_CELL_TAG) && reader.isStartElement();
}

bool ImportOds::isCellEnd(const QXmlStreamReader& reader) const
{
    return (reader.name().toString() == TABLE_CELL_TAG) &&
           reader.isEndElement();
}

ColumnType ImportOds::recognizeColumnType(ColumnType currentType,
                                          const QString& xmlColTypeValue) const
{
    if (xmlColTypeValue == STRING_TAG)
    {
        if (currentType == ColumnType::UNKNOWN)
            return ColumnType::STRING;
        if (currentType != ColumnType::STRING)
            return ColumnType::STRING;
    }

    if (xmlColTypeValue == DATE_TAG)
    {
        if (currentType == ColumnType::UNKNOWN)
            return ColumnType::DATE;
        if (currentType != ColumnType::DATE)
            return ColumnType::STRING;
    }

    if ((xmlColTypeValue == FLOAT_TAG) || (xmlColTypeValue == PERCENTAGE_TAG) ||
        (xmlColTypeValue == CURRENCY_TAG) || (xmlColTypeValue == TIME_TAG))
    {
        if (currentType == ColumnType::UNKNOWN)
            return ColumnType::NUMBER;
        if (currentType != ColumnType::NUMBER)
            return ColumnType::STRING;
    }

    return currentType;
}

QVariant ImportOds::retrieveValueFromStringColumn(
    QXmlStreamReader& reader) const
{
    const QString xmlColTypeValue{
        reader.attributes().value(OFFICE_VALUE_TYPE_TAG).toString()};
    const QString emptyString(QString::fromLatin1(""));
    const QString currentDateValue{
        reader.attributes().value(OFFICE_DATE_VALUE_TAG).toString()};

    while ((!reader.atEnd()) && (reader.name() != P_TAG))
        reader.readNext();

    while ((reader.tokenType() != QXmlStreamReader::Characters) &&
           (reader.name() != TABLE_CELL_TAG))
        reader.readNext();

    if (xmlColTypeValue == DATE_TAG)
        return {currentDateValue};

    const QStringView stringView{reader.text()};
    if (stringView.isNull())
        return {emptyString};

    return stringView.toString();
}
QVariant ImportOds::retrieveValueFromField(QXmlStreamReader& reader,
                                           ColumnType columnType) const
{
    Q_ASSERT(columnType != ColumnType::UNKNOWN);

    if (columnType == ColumnType::STRING)
        return retrieveValueFromStringColumn(reader);

    if (columnType == ColumnType::NUMBER)
        return {reader.attributes().value(OFFICE_VALUE_TAG).toDouble()};

    QString dateValue{
        reader.attributes().value(OFFICE_DATE_VALUE_TAG).toString()};
    dateValue.chop(dateValue.size() - ODS_STRING_DATE_LENGTH);
    return QDate::fromString(dateValue, DATE_FORMAT);
}

bool ImportOds::isOfficeValueTagEmpty(const QXmlStreamReader& reader) const
{
    const QString xmlColTypeValue{
        reader.attributes().value(OFFICE_VALUE_TYPE_TAG).toString()};
    return xmlColTypeValue.isEmpty();
}

bool ImportOds::isSheetAvailable(const QString& sheetName)
{
    if ((!sheetNames_) && (!getSheetNames().first))
        return false;

    return isSheetNameValid(*sheetNames_, sheetName);
}

bool ImportOds::initializeColumnTypes(const QString& sheetName)
{
    QuaZipFile quaZipFile;
    if (!initZipFile(quaZipFile, QStringLiteral("content.xml")))
        return false;

    QXmlStreamReader reader;
    moveToSecondRow(sheetName, quaZipFile, reader);

    return analyzeSheet(sheetName);
}
