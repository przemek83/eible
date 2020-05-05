#include "ImportOds.h"

#include <utility>

#include <quazip5/quazip.h>
#include <quazip5/quazipfile.h>
#include <QVariant>
#include <QVector>
#include <QXmlStreamReader>
#include <QtXml/QDomDocument>

#include "EibleUtilities.h"

const QString ImportOds::TABLE_TAG{QStringLiteral("table")};
const QString ImportOds::TABLE_ROW_TAG{QStringLiteral("table-row")};
const QString ImportOds::TABLE_CELL_TAG{QStringLiteral("table-cell")};
const QString ImportOds::OFFICE_VALUE_TYPE_TAG{
    QStringLiteral("office:value-type")};
const QString ImportOds::COLUMNS_REPEATED_TAG{
    QStringLiteral("table:number-columns-repeated")};
const QString ImportOds::STRING_TAG{QStringLiteral("string")};
const QString ImportOds::DATE_TAG{QStringLiteral("date")};
const QString ImportOds::FLOAT_TAG{QStringLiteral("float")};
const QString ImportOds::PERCENTAGE_TAG{QStringLiteral("percentage")};
const QString ImportOds::CURRENCY_TAG{QStringLiteral("currency")};
const QString ImportOds::TIME_TAG{QStringLiteral("time")};
const QString ImportOds::P_TAG{QStringLiteral("p")};
const QString ImportOds::OFFICE_DATE_VALUE_TAG{
    QStringLiteral("office:date-value")};
const QString ImportOds::OFFICE_VALUE_TAG{QStringLiteral("office:value")};
const QString ImportOds::DATE_FORMAT{QStringLiteral("yyyy-MM-dd")};
const QString ImportOds::TABLE_QUALIFIED_NAME{"table:table"};
const QString ImportOds::TABLE_NAME_TAG{QLatin1String("table:name")};

ImportOds::ImportOds(QIODevice& ioDevice)
    : ImportSpreadsheet(ioDevice), zip_(&ioDevice_)
{
}

std::pair<bool, QStringList> ImportOds::getSheetNames()
{
    if (sheetNames_)
        return {true, *sheetNames_};

    QuaZipFile zipFile;
    if (!openZipFile(zipFile, QStringLiteral("settings.xml")))
        return {false, {}};

    // Create, set content and read DOM.
    QDomDocument xmlDocument(__FUNCTION__);
    if (!xmlDocument.setContent(zipFile.readAll()))
    {
        setError(__FUNCTION__, "Xml file is damaged.");
        return {false, {}};
    }

    QDomElement root = xmlDocument.documentElement();

    const QString configMapNamed(
        QStringLiteral("config:config-item-map-named"));
    const QString configName(QStringLiteral("config:name"));
    const QString configMapEntry(
        QStringLiteral("config:config-item-map-entry"));
    const QString tables(QStringLiteral("Tables"));

    int elementsCount = root.elementsByTagName(configMapNamed).size();
    QStringList sheetNames{};
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

    sheetNames_ = std::move(sheetNames);
    return {true, *sheetNames_};
}

std::pair<bool, QVector<ColumnType>> ImportOds::getColumnTypes(
    const QString& sheetName)
{
    if (const auto it = columnTypes_.find(sheetName); it != columnTypes_.end())
        return {true, *it};

    if (!sheetNames_ && !getSheetNames().first)
        return {false, {}};

    QuaZipFile zipFile;
    if (!openZipFile(zipFile, QStringLiteral("content.xml")))
        return {false, {}};

    QXmlStreamReader xmlStreamReader;
    if (!moveToSecondRow(sheetName, zipFile, xmlStreamReader))
        return {false, {}};

    QVector<ColumnType> columnTypes;

    // Actual column number.
    int column = NOT_SET_COLUMN;
    int maxColumn = NOT_SET_COLUMN;

    // Actual row number.
    int rowCounter = 0;

    // Actual data type in current cell (s, str, null).
    QString currentColType;

    bool rowEmpty = true;

    while (!xmlStreamReader.atEnd() &&
           xmlStreamReader.name().compare(TABLE_TAG) != 0)
    {
        // If start of row encountered than reset column counter add increment
        // row counter.
        if (isRowStart(xmlStreamReader))
            column = NOT_SET_COLUMN;

        // When we encounter start of cell description.
        if (isCellStart(xmlStreamReader))
        {
            column++;

            // Remember column type.
            currentColType = xmlStreamReader.attributes()
                                 .value(OFFICE_VALUE_TYPE_TAG)
                                 .toString();

            const int repeats{
                getColumnRepeatCount(xmlStreamReader.attributes())};

            for (int i = 0; i < repeats; ++i)
            {
                if (0 == currentColType.compare(STRING_TAG))
                {
                    maxColumn = std::max(maxColumn, column + i);
                    rowEmpty = false;
                    while (column + i >= columnTypes.size())
                        columnTypes.push_back(ColumnType::UNKNOWN);
                    if (columnTypes.at(column + i) == ColumnType::UNKNOWN)
                    {
                        columnTypes[column + i] = ColumnType::STRING;
                    }
                    else
                    {
                        if (columnTypes.at(column + i) != ColumnType::STRING)
                        {
                            columnTypes[column + i] = ColumnType::STRING;
                        }
                    }
                }
                else
                {
                    if (0 == currentColType.compare(DATE_TAG))
                    {
                        maxColumn = std::max(maxColumn, column + i);
                        rowEmpty = false;
                        while (column + i >= columnTypes.size())
                            columnTypes.push_back(ColumnType::UNKNOWN);
                        if (columnTypes.at(column + i) == ColumnType::UNKNOWN)
                        {
                            columnTypes[column + i] = ColumnType::DATE;
                        }
                        else
                        {
                            if (columnTypes.at(column + i) != ColumnType::DATE)
                            {
                                columnTypes[column + i] = ColumnType::STRING;
                            }
                        }
                    }
                    else
                    {
                        if (0 == currentColType.compare(FLOAT_TAG) ||
                            0 == currentColType.compare(PERCENTAGE_TAG) ||
                            0 == currentColType.compare(CURRENCY_TAG) ||
                            0 == currentColType.compare(TIME_TAG))
                        {
                            maxColumn = std::max(maxColumn, column + i);
                            rowEmpty = false;
                            while (column + i >= columnTypes.size())
                                columnTypes.push_back(ColumnType::UNKNOWN);
                            if (columnTypes.at(column + i) ==
                                ColumnType::UNKNOWN)
                            {
                                columnTypes[column + i] = ColumnType::NUMBER;
                            }
                            else
                            {
                                if (columnTypes.at(column + i) !=
                                    ColumnType::NUMBER)
                                {
                                    columnTypes[column + i] =
                                        ColumnType::STRING;
                                }
                            }
                        }
                    }
                }
            }
            column += repeats - 1;
        }
        xmlStreamReader.readNext();
        if (0 == xmlStreamReader.name().compare(TABLE_ROW_TAG) &&
            xmlStreamReader.isEndElement())
        {
            if (!rowEmpty)
                rowCounter++;
            rowEmpty = true;
        }
    }

    for (int i = 0; i < columnTypes.size(); ++i)
    {
        if (ColumnType::UNKNOWN == columnTypes.at(i))
            columnTypes[i] = ColumnType::STRING;
    }

    rowCounts_[sheetName] = rowCounter;
    columnCounts_[sheetName] = maxColumn + 1;
    columnTypes_[sheetName] = columnTypes;

    return {true, columnTypes};
}

std::pair<bool, QStringList> ImportOds::getColumnNames(const QString& sheetName)
{
    if (!sheetNames_ && !getSheetNames().first)
        return {false, {}};

    if (!sheetNameValid(sheetName))
        return {false, {}};

    const auto it = columnCounts_.find(sheetName);
    if (it == columnCounts_.end() && !analyzeSheet(sheetName))
        return {false, {}};

    QuaZipFile zipFile;
    if (!openZipFile(zipFile, QStringLiteral("content.xml")))
        return {false, {}};

    QXmlStreamReader xmlStreamReader;
    xmlStreamReader.setDevice(&zipFile);
    skipToSheet(xmlStreamReader, sheetName);

    while (!xmlStreamReader.atEnd() && xmlStreamReader.name() != TABLE_ROW_TAG)
        xmlStreamReader.readNext();

    xmlStreamReader.readNext();

    // Actual column number.
    int column = NOT_SET_COLUMN;
    QXmlStreamReader::TokenType lastToken = xmlStreamReader.tokenType();
    QStringList columnNames{};

    // Parse first row.
    while (!xmlStreamReader.atEnd() && xmlStreamReader.name() != TABLE_ROW_TAG)
    {
        // When we encounter first cell of worksheet.
        if (isCellStart(xmlStreamReader))
        {
            QString emptyColNumber = xmlStreamReader.attributes()
                                         .value(COLUMNS_REPEATED_TAG)
                                         .toString();
            if (!emptyColNumber.isEmpty())
                break;

            // Add column number.
            column++;
        }

        // If we encounter start of cell content we add it to list.
        if (!xmlStreamReader.atEnd() &&
            xmlStreamReader.name().toString() == P_TAG &&
            xmlStreamReader.tokenType() == QXmlStreamReader::StartElement)
        {
            while (xmlStreamReader.tokenType() != QXmlStreamReader::Characters)
                xmlStreamReader.readNext();

            columnNames.push_back(xmlStreamReader.text().toString());
        }

        // If we encounter empty cell we add it to list.
        if (isCellEnd(xmlStreamReader) &&
            lastToken == QXmlStreamReader::StartElement)
            columnNames << emptyColName_;

        lastToken = xmlStreamReader.tokenType();
        xmlStreamReader.readNext();
    }

    const unsigned int columnCount{columnCounts_[sheetName]};
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

std::pair<bool, QVector<QVector<QVariant>>> ImportOds::getLimitedData(
    const QString& sheetName, const QVector<unsigned int>& excludedColumns,
    unsigned int rowLimit)
{
    const auto [success, columnTypes] = getColumnTypes(sheetName);
    if (!success)
        return {false, {}};

    const unsigned int columnCount{columnCounts_[sheetName]};
    if (!columnsToExcludeAreValid(excludedColumns, columnCount))
        return {false, {}};

    QuaZipFile zipFile;
    if (!openZipFile(zipFile, QStringLiteral("content.xml")))
        return {false, {}};

    QXmlStreamReader xmlStreamReader;
    if (!moveToSecondRow(sheetName, zipFile, xmlStreamReader))
        return {false, {}};

    QVector<QVariant> templateDataRow{
        createTemplateDataRow(excludedColumns, columnTypes)};

    QMap<unsigned int, unsigned int> activeColumnsMapping{
        createActiveColumnMapping(excludedColumns, columnCount)};

    QVector<QVector<QVariant>> dataContainer(
        std::min(getRowCount(sheetName).second, rowLimit));

    // Current row data.
    QVector<QVariant> currentDataRow(templateDataRow);

    // Actual column number.
    int column = NOT_SET_COLUMN;

    // Actual data type in cell (s, str, null).
    QString currentColType;
    QString currentDateValue{};

    // Current row number.
    unsigned int rowCounter = 0;
    int cellsFilledInRow = 0;

    const QString emptyString(QLatin1String(""));
    bool rowEmpty = true;
    unsigned int lastEmittedPercent{0};
    const unsigned int rowCount{rowCounts_[sheetName]};
    const bool fillSamplesOnly{rowCount != rowLimit};
    while (!xmlStreamReader.atEnd() &&
           0 != xmlStreamReader.name().compare(TABLE_TAG) &&
           rowCounter < rowLimit)
    {
        // If start of row encountered than reset column counter add
        // increment row counter.
        if (isRowStart(xmlStreamReader))
        {
            column = NOT_SET_COLUMN;
            if (!rowEmpty)
            {
                dataContainer[rowCounter] = currentDataRow;
                currentDataRow = templateDataRow;
                rowCounter++;
                updateProgress(rowCounter, rowLimit, lastEmittedPercent);
            }
            rowEmpty = true;

            if (fillSamplesOnly &&
                rowCounter > static_cast<unsigned int>(dataContainer.size()))
                break;
        }

        // When we encounter start of cell description.
        if (isCellStart(xmlStreamReader))
        {
            column++;

            // If we encounter column outside expected grid we move to row end.
            if (column >= static_cast<int>(columnCount))
            {
                while (!xmlStreamReader.atEnd() && !isRowEnd(xmlStreamReader))
                    xmlStreamReader.readNext();
                continue;
            }

            // Remember cell type.
            currentColType = xmlStreamReader.attributes()
                                 .value(OFFICE_VALUE_TYPE_TAG)
                                 .toString();

            int repeats{getColumnRepeatCount(xmlStreamReader.attributes())};
            if (column + repeats - 1 >= static_cast<int>(columnCount))
                repeats = columnCount - column;

            if (!currentColType.isEmpty())
            {
                QVariant value;
                ColumnType format = columnTypes.at(column);

                switch (format)
                {
                    case ColumnType::STRING:
                    {
                        currentDateValue = xmlStreamReader.attributes()
                                               .value(OFFICE_DATE_VALUE_TAG)
                                               .toString();

                        while (!xmlStreamReader.atEnd() &&
                               0 != xmlStreamReader.name().compare(P_TAG))

                        {
                            xmlStreamReader.readNext();
                        }

                        while (
                            xmlStreamReader.tokenType() !=
                                QXmlStreamReader::Characters &&
                            0 != xmlStreamReader.name().compare(TABLE_CELL_TAG))
                        {
                            xmlStreamReader.readNext();
                        }
                        rowEmpty = false;
                        if (currentColType.compare(DATE_TAG) == 0)
                        {
                            value = QVariant(currentDateValue);
                        }
                        else
                        {
                            const QString* stringPointer =
                                xmlStreamReader.text().string();
                            value = QVariant(stringPointer == nullptr
                                                 ? emptyString
                                                 : *stringPointer);
                        }
                        break;
                    }

                    case ColumnType::DATE:
                    {
                        static const int odsStringDateLength{10};
                        rowEmpty = false;
                        value = QVariant(
                            QDate::fromString(xmlStreamReader.attributes()
                                                  .value(OFFICE_DATE_VALUE_TAG)
                                                  .toString()
                                                  .left(odsStringDateLength),
                                              DATE_FORMAT));

                        break;
                    }

                    case ColumnType::NUMBER:
                    {
                        rowEmpty = false;
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

                for (int i = 0; i < repeats; ++i)
                {
                    if (!fillSamplesOnly &&
                        excludedColumns.contains(column + i))
                        continue;

                    cellsFilledInRow++;
                    currentDataRow[activeColumnsMapping[column + i]] = value;
                }
            }

            column += repeats - 1;
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

    if (!sheetNameValid(sheetName))
        return {false, {}};

    if (const auto it = countMap.find(sheetName); it != countMap.end())
        return {true, it.value()};

    if (!analyzeSheet(sheetName))
        return {false, 0};

    return {true, countMap.find(sheetName).value()};
}

bool ImportOds::analyzeSheet(const QString& sheetName)
{
    QuaZipFile zipFile;
    if (!openZipFile(zipFile, QStringLiteral("content.xml")))
        return false;

    QXmlStreamReader xmlStreamReader;
    if (!moveToSecondRow(sheetName, zipFile, xmlStreamReader))
        return false;

    std::tie(rowCounts_[sheetName], columnCounts_[sheetName]) =
        getRowAndColumnCount(xmlStreamReader);
    return true;
}

std::pair<unsigned int, unsigned int> ImportOds::getRowAndColumnCount(
    QXmlStreamReader& xmlStreamReader) const
{
    int column = NOT_SET_COLUMN;
    int rowCounter = 0;
    int maxColumn = NOT_SET_COLUMN;
    bool rowEmpty = true;

    while (!xmlStreamReader.atEnd() &&
           xmlStreamReader.name().compare(TABLE_TAG) != 0)
    {
        if (isRowStart(xmlStreamReader))
            column = NOT_SET_COLUMN;

        if (isCellStart(xmlStreamReader))
        {
            const QXmlStreamAttributes attributes{xmlStreamReader.attributes()};
            column += getColumnRepeatCount(attributes);
            if (isRecognizedColumnType(attributes))
            {
                maxColumn = std::max(maxColumn, column);
                rowEmpty = false;
            }
        }
        xmlStreamReader.readNext();
        if (isRowEnd(xmlStreamReader))
        {
            if (!rowEmpty)
                rowCounter++;
            rowEmpty = true;
        }
    }
    return {rowCounter, maxColumn + 1};
}

bool ImportOds::moveToSecondRow(const QString& sheetName, QuaZipFile& zipFile,
                                QXmlStreamReader& xmlStreamReader)
{
    xmlStreamReader.setDevice(&zipFile);
    skipToSheet(xmlStreamReader, sheetName);

    bool secondRow{false};
    while (!xmlStreamReader.atEnd() &&
           !(xmlStreamReader.qualifiedName() == TABLE_QUALIFIED_NAME &&
             xmlStreamReader.tokenType() == QXmlStreamReader::EndElement))
    {
        if (xmlStreamReader.name() == TABLE_ROW_TAG &&
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

bool ImportOds::openZipFile(QuaZipFile& zipFile, const QString& zipFileName)
{
    if (!zip_.isOpen() && !zip_.open(QuaZip::mdUnzip))
    {
        setError(__FUNCTION__,
                 "Can not open zip file " + zip_.getZipName() + ".");
        return false;
    }

    if (!zip_.setCurrentFile(zipFileName))
    {
        setError(__FUNCTION__,
                 "Can not find file " + zipFileName + " in archive.");
        return false;
    }

    zipFile.setZip(&zip_);
    if (!zipFile.open(QIODevice::ReadOnly | QIODevice::Text))
    {
        setError(__FUNCTION__,
                 "Can not open file " + zipFile.getFileName() + ".");
        return false;
    }

    return true;
}

bool ImportOds::isRecognizedColumnType(
    const QXmlStreamAttributes& attributes) const
{
    const QString columnType =
        attributes.value(OFFICE_VALUE_TYPE_TAG).toString();

    return 0 == columnType.compare(STRING_TAG) ||
           0 == columnType.compare(DATE_TAG) ||
           0 == columnType.compare(FLOAT_TAG) ||
           0 == columnType.compare(PERCENTAGE_TAG) ||
           0 == columnType.compare(CURRENCY_TAG) ||
           0 == columnType.compare(TIME_TAG);
}

int ImportOds::getColumnRepeatCount(
    const QXmlStreamAttributes& attributes) const
{
    return std::max(1,
                    attributes.value(COLUMNS_REPEATED_TAG).toString().toInt());
}

bool ImportOds::isRowStart(const QXmlStreamReader& xmlStreamReader) const
{
    return 0 == xmlStreamReader.name().compare(TABLE_ROW_TAG) &&
           xmlStreamReader.isStartElement();
}

bool ImportOds::isRowEnd(const QXmlStreamReader& xmlStreamReader) const
{
    return 0 == xmlStreamReader.name().compare(TABLE_ROW_TAG) &&
           xmlStreamReader.isEndElement();
}

bool ImportOds::isCellStart(const QXmlStreamReader& xmlStreamReader) const
{
    return 0 == xmlStreamReader.name().compare(TABLE_CELL_TAG) &&
           xmlStreamReader.isStartElement();
}

bool ImportOds::isCellEnd(const QXmlStreamReader& xmlStreamReader) const
{
    return xmlStreamReader.name().toString() == TABLE_CELL_TAG &&
           xmlStreamReader.isEndElement();
}

bool ImportOds::sheetNameValid(const QString& sheetName)
{
    if (!sheetNames_->contains(sheetName))
    {
        setError(__FUNCTION__,
                 "Sheet " + sheetName +
                     " not found. Available sheets: " + sheetNames_->join(','));
        return false;
    }
    return true;
}

bool ImportOds::columnsToExcludeAreValid(
    const QVector<unsigned int>& excludedColumns, unsigned int columnCount)
{
    auto it = std::find_if(
        excludedColumns.begin(), excludedColumns.end(),
        [=](unsigned int column) { return column >= columnCount; });
    if (it != excludedColumns.end())
    {
        setError(__FUNCTION__, "Column to exclude " + QString::number(*it) +
                                   " is invalid. Xlsx got only " +
                                   QString::number(columnCount) +
                                   " columns indexed from 0.");
        return false;
    }
    return true;
}

QVector<QVariant> ImportOds::createTemplateDataRow(
    const QVector<unsigned int>& excludedColumns,
    const QVector<ColumnType>& columnTypes) const
{
    QVector<QVariant> templateDataRow;
    int columnToFill = 0;
    templateDataRow.resize(columnTypes.size() - excludedColumns.size());
    for (int i = 0; i < columnTypes.size(); ++i)
    {
        if (!excludedColumns.contains(i))
        {
            templateDataRow[columnToFill] =
                EibleUtilities::getDefaultVariantForFormat(columnTypes[i]);
            columnToFill++;
        }
    }
    return templateDataRow;
}

QMap<unsigned int, unsigned int> ImportOds::createActiveColumnMapping(
    const QVector<unsigned int>& excludedColumns,
    unsigned int columnCount) const
{
    QMap<unsigned int, unsigned int> activeColumnsMapping;
    int columnToFill = 0;
    for (unsigned int i = 0; i < columnCount; ++i)
    {
        if (!excludedColumns.contains(i))
        {
            activeColumnsMapping[i] = columnToFill;
            columnToFill++;
        }
    }
    return activeColumnsMapping;
}
