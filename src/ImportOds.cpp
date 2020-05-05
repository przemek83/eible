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

    int repeatCount = 1;

    bool rowEmpty = true;

    while (!xmlStreamReader.atEnd() &&
           xmlStreamReader.name().compare(TABLE_TAG) != 0)
    {
        // If start of row encountered than reset column counter add increment
        // row counter.
        if (0 == xmlStreamReader.name().compare(TABLE_ROW_TAG) &&
            xmlStreamReader.isStartElement())
            column = NOT_SET_COLUMN;

        // When we encounter start of cell description.
        if (0 == xmlStreamReader.name().compare(TABLE_CELL_TAG) &&
            xmlStreamReader.tokenType() == QXmlStreamReader::StartElement)
        {
            column++;

            // Remember column type.
            currentColType = xmlStreamReader.attributes()
                                 .value(OFFICE_VALUE_TYPE_TAG)
                                 .toString();

            // Number of repeats.
            repeatCount = xmlStreamReader.attributes()
                              .value(COLUMNS_REPEATED_TAG)
                              .toString()
                              .toInt();

            if (0 == repeatCount)
                repeatCount = 1;

            for (int i = 0; i < repeatCount; ++i)
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
            column += repeatCount - 1;
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
    const unsigned int columnCount{columnCounts_[sheetName]};

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
        if (xmlStreamReader.name().toString() == TABLE_CELL_TAG &&
            xmlStreamReader.tokenType() == QXmlStreamReader::StartElement)
        {
            QString emptyColNumber = xmlStreamReader.attributes()
                                         .value(COLUMNS_REPEATED_TAG)
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
            xmlStreamReader.name().toString() == P_TAG &&
            xmlStreamReader.tokenType() == QXmlStreamReader::StartElement)
        {
            while (xmlStreamReader.tokenType() != QXmlStreamReader::Characters)
            {
                xmlStreamReader.readNext();
            }

            columnNames.push_back(xmlStreamReader.text().toString());
        }

        // If we encounter empty cell we add it to list.
        if (xmlStreamReader.name().toString() == TABLE_CELL_TAG &&
            xmlStreamReader.tokenType() == QXmlStreamReader::EndElement &&
            lastToken == QXmlStreamReader::StartElement)
        {
            columnNames << emptyColName_;
        }

        lastToken = xmlStreamReader.tokenType();
        xmlStreamReader.readNext();
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

std::pair<bool, QVector<QVector<QVariant>>> ImportOds::getLimitedData(
    const QString& sheetName, const QVector<unsigned int>& excludedColumns,
    unsigned int rowLimit)
{
    const auto [success, columnTypes] = getColumnTypes(sheetName);
    if (!success)
        return {false, {}};

    const unsigned int columnCount{columnCounts_[sheetName]};
    const unsigned int rowCount{rowCounts_[sheetName]};
    const bool fillSamplesOnly{rowCount != rowLimit};

    auto it = std::find_if(
        excludedColumns.begin(), excludedColumns.end(),
        [=](unsigned int column) { return column >= columnCount; });
    if (it != excludedColumns.end())
    {
        setError(__FUNCTION__, "Column to exclude " + QString::number(*it) +
                                   " is invalid. Xlsx got only " +
                                   QString::number(columnCount) +
                                   " columns indexed from 0.");
        return {false, {}};
    }

    QuaZipFile zipFile;
    if (!openZipFile(zipFile, QStringLiteral("content.xml")))
        return {false, {}};

    QXmlStreamReader xmlStreamReader;
    if (!moveToSecondRow(sheetName, zipFile, xmlStreamReader))
        return {false, {}};

    // Null elements row.
    QVector<QVariant> templateDataRow;

    QMap<int, int> activeColumnsMapping;

    int columnToFill = 0;

    templateDataRow.resize(
        fillSamplesOnly ? columnCount : columnCount - excludedColumns.size());
    for (unsigned int i = 0; i < columnCount; ++i)
    {
        if (fillSamplesOnly || !excludedColumns.contains(i))
        {
            templateDataRow[columnToFill] =
                EibleUtilities::getDefaultVariantForFormat(columnTypes[i]);
            activeColumnsMapping[i] = columnToFill;
            columnToFill++;
        }
    }

    QVector<QVector<QVariant>> dataContainer(
        std::min(getRowCount(sheetName).second, rowLimit));

    // Protection from potential core related to empty rows.
    unsigned int containerSize = dataContainer.size();
    for (unsigned int i = 0; i < containerSize; ++i)
        dataContainer[i] = templateDataRow;

    // Current row data.
    QVector<QVariant> currentDataRow(templateDataRow);

    // Actual column number.
    int column = NOT_SET_COLUMN;

    // Actual data type in cell (s, str, null).
    QString currentColType;
    QString currentDateValue{};

    // Current row number.
    unsigned int rowCounter = 0;

    int repeatCount = 1;

    int cellsFilledInRow = 0;

    const QString emptyString(QLatin1String(""));
    bool rowEmpty = true;
    unsigned int lastEmittedPercent{0};
    while (!xmlStreamReader.atEnd() &&
           0 != xmlStreamReader.name().compare(TABLE_TAG) &&
           rowCounter < rowLimit)
    {
        // If start of row encountered than reset column counter add
        // increment row counter.
        if (0 == xmlStreamReader.name().compare(TABLE_ROW_TAG) &&
            xmlStreamReader.isStartElement())
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

            if (fillSamplesOnly && rowCounter > containerSize)
                break;
        }

        // When we encounter start of cell description.
        if (0 == xmlStreamReader.name().compare(TABLE_CELL_TAG) &&
            xmlStreamReader.isStartElement())
        {
            column++;

            // If we encounter column outside expected grid we move to row end.
            if (column >= static_cast<int>(columnCount))
            {
                while (!xmlStreamReader.atEnd() &&
                       !(0 == xmlStreamReader.name().compare(TABLE_ROW_TAG) &&
                         xmlStreamReader.isEndElement()))
                {
                    xmlStreamReader.readNext();
                }
                continue;
            }

            // Remember cell type.
            currentColType = xmlStreamReader.attributes()
                                 .value(OFFICE_VALUE_TYPE_TAG)
                                 .toString();

            // Number of repeats.
            repeatCount = xmlStreamReader.attributes()
                              .value(COLUMNS_REPEATED_TAG)
                              .toString()
                              .toInt();

            if (0 == repeatCount)
            {
                repeatCount = 1;
            }

            if (column + repeatCount - 1 >= static_cast<int>(columnCount))
            {
                repeatCount = columnCount - column;
            }

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

                for (int i = 0; i < repeatCount; ++i)
                {
                    if (!fillSamplesOnly &&
                        excludedColumns.contains(column + i))
                    {
                        continue;
                    }

                    cellsFilledInRow++;
                    currentDataRow[activeColumnsMapping[column + i]] = value;
                }
            }

            column += repeatCount - 1;
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
    QuaZipFile zipFile;
    if (!openZipFile(zipFile, QStringLiteral("content.xml")))
        return false;

    QXmlStreamReader xmlStreamReader;
    if (!moveToSecondRow(sheetName, zipFile, xmlStreamReader))
        return false;

    int column = NOT_SET_COLUMN;
    int rowCounter = 0;
    int repeatCount = 1;
    int maxColumn = NOT_SET_COLUMN;
    QString currentColType;
    bool rowEmpty = true;

    while (!xmlStreamReader.atEnd() &&
           xmlStreamReader.name().compare(TABLE_TAG) != 0)
    {
        if (0 == xmlStreamReader.name().compare(TABLE_ROW_TAG) &&
            xmlStreamReader.isStartElement())
            column = NOT_SET_COLUMN;

        if (0 == xmlStreamReader.name().compare(TABLE_CELL_TAG) &&
            xmlStreamReader.isStartElement())
        {
            column++;

            currentColType = xmlStreamReader.attributes()
                                 .value(OFFICE_VALUE_TYPE_TAG)
                                 .toString();

            repeatCount = std::max(1, xmlStreamReader.attributes()
                                          .value(COLUMNS_REPEATED_TAG)
                                          .toString()
                                          .toInt());

            for (int i = 0; i < repeatCount; ++i)
            {
                if (0 == currentColType.compare(STRING_TAG) ||
                    0 == currentColType.compare(DATE_TAG) ||
                    0 == currentColType.compare(FLOAT_TAG) ||
                    0 == currentColType.compare(PERCENTAGE_TAG) ||
                    0 == currentColType.compare(CURRENCY_TAG) ||
                    0 == currentColType.compare(TIME_TAG))
                {
                    maxColumn = std::max(maxColumn, column + i);
                    rowEmpty = false;
                }
            }
            column += repeatCount - 1;
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

    rowCounts_[sheetName] = rowCounter;
    columnCounts_[sheetName] = maxColumn + 1;

    return true;
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
