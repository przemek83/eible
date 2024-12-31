#include <eible/ImportXlsx.h>

#include <algorithm>
#include <cmath>

#include <quazip/quazip.h>
#include <quazip/quazipfile.h>
#include <QIODevice>
#include <QMap>
#include <QSet>
#include <QVariant>
#include <QXmlStreamReader>
#include <QtXml/QDomDocument>

#include "Utilities.h"

ImportXlsx::ImportXlsx(QIODevice& ioDevice)
    : ImportSpreadsheet(ioDevice),
      excelColNames_(
          utilities::generateExcelColumnNames(utilities::getMaxExcelColumns()))
{
}

std::pair<bool, QStringList> ImportXlsx::getSheetNames()
{
    if (sheets_)
    {
        QStringList sheets;
        sheets.reserve(sheets_->size());
        for (const auto& [sheetName, sheetPath] : *sheets_)
            sheets << sheetName;
        return {true, sheets};
    }

    auto [success,
          sheetIdToUserFriendlyNameMap]{getSheetIdToUserFriendlyNameMap()};
    if (!success)
        return {false, {}};

    std::tie(success, sheets_) = retrieveSheets(sheetIdToUserFriendlyNameMap);
    if (!success)
        return {false, {}};

    QStringList sheetsToReturn;
    sheetsToReturn.reserve(sheets_->size());
    for (const auto& [sheetName, sheetPath] : *sheets_)
        sheetsToReturn << sheetName;
    return {true, sheetsToReturn};
}

std::pair<bool, QStringList> ImportXlsx::getColumnNames(
    const QString& sheetName)
{
    if (!isCommonDataOk())
        return {false, {}};

    if (!isSheetNameValid(getSheetNames().second, sheetName))
        return {false, {}};

    if (const auto it{columnNames_.constFind(sheetName)};
        (it == columnNames_.constEnd()) && (!analyzeSheet(sheetName)))
        return {false, {}};

    return {true, columnNames_.value(sheetName)};
}

QList<int> ImportXlsx::retrieveDateStyles(const QDomNodeList& sheetNodes)
{
    QList<int> dateStyles;
    const QList predefinedExcelStylesForDates{14, 15, 16, 17, 22};
    dateStyles.append(predefinedExcelStylesForDates);

    const int sheetsNodesCount{sheetNodes.size()};
    for (int i{0}; i < sheetsNodesCount; ++i)
    {
        const QDomElement sheet{sheetNodes.at(i).toElement()};
        if (sheet.hasAttribute(QStringLiteral("numFmtId")) &&
            sheet.hasAttribute(QStringLiteral("formatCode")))
        {
            const QString formatCode{
                sheet.attribute(QStringLiteral("formatCode"))};
            const bool gotD{
                formatCode.contains(QStringLiteral("d"), Qt::CaseInsensitive)};
            const bool gotM{
                formatCode.contains(QStringLiteral("m"), Qt::CaseInsensitive)};
            const bool gotY{
                formatCode.contains(QStringLiteral("y"), Qt::CaseInsensitive)};

            const bool dayAndYearFormat{gotD && gotY};
            const bool dayAndMonthFormat{gotD && gotM};
            const bool monthAndYearFormat{gotM && gotY};

            if (dayAndYearFormat || dayAndMonthFormat || monthAndYearFormat)
                dateStyles.push_back(
                    sheet.attribute(QStringLiteral("numFmtId")).toInt());
        }
    }
    return dateStyles;
}

QList<int> ImportXlsx::retrieveAllStyles(const QDomNodeList& sheetNodes)
{
    QList<int> allStyles;

    const int sheetsNodesCount{sheetNodes.size()};
    allStyles.reserve(sheetsNodesCount);
    const QString searchedAttribute{QStringLiteral("numFmtId")};
    for (int i{0}; i < sheetsNodesCount; ++i)
    {
        const QDomElement sheet{sheetNodes.at(i).toElement()};
        if ((!sheet.isNull()) && sheet.hasAttribute(searchedAttribute))
            allStyles.push_back(sheet.attribute(searchedAttribute).toInt());
    }
    return allStyles;
}

std::tuple<bool, std::optional<QList<int>>, std::optional<QList<int>>>
ImportXlsx::getStyles()
{
    QuaZipFile zipFile;
    if (!initZipFile(zipFile, QStringLiteral("xl/styles.xml")))
        return {false, {}, {}};

    QDomDocument xmlDocument;
    if (!xmlDocument.setContent(zipFile.readAll()))
    {
        setError(QStringLiteral("Xml file is corrupted."));
        return {false, std::nullopt, std::nullopt};
    }

    const QDomElement root{xmlDocument.documentElement()};
    QDomNodeList sheetNodes{
        root.firstChildElement(QStringLiteral("numFmts")).childNodes()};
    const QList<int> dateStyles{retrieveDateStyles(sheetNodes)};

    sheetNodes = root.firstChildElement(QStringLiteral("cellXfs")).childNodes();
    const QList<int> allStyles{retrieveAllStyles(sheetNodes)};

    return {true, dateStyles, allStyles};
}

std::pair<bool, QString> ImportXlsx::getSheetPath(const QString& sheetName)
{
    for (const auto& [currentSheetName, sheetPath] : *sheets_)
        if (currentSheetName == sheetName)
            return {true, sheetPath};

    QStringList sheetNames;
    sheetNames.reserve(sheets_->size());
    for (const auto& [currenSheetName, sheetPath] : *sheets_)
        sheetNames << currenSheetName;
    setError("Can not find sheet path for sheet name " + sheetName +
             ". Available sheet names:" + sheetNames.join(','));
    return {false, {}};
}

std::pair<bool, QStringList> ImportXlsx::getSharedStrings()
{
    if (sharedStrings_)
        return {true, *sharedStrings_};

    if (!openZip())
        return {false, {}};

    // There can be no shared strings file. Return empty list in that case.
    if (!setCurrentZipFile(QStringLiteral("xl/sharedStrings.xml")))
        return {true, {}};

    QuaZipFile zipFile;
    if (!openZipFile(zipFile))
        return {false, {}};

    QXmlStreamReader xmlStreamReader;
    xmlStreamReader.setDevice(&zipFile);

    QStringList sharedStrings;
    while (!xmlStreamReader.atEnd())
    {
        xmlStreamReader.readNext();
        if (xmlStreamReader.name() == T_TAG)
            sharedStrings.append(xmlStreamReader.readElementText());
    }

    sharedStrings_ = std::move(sharedStrings);
    return {true, *sharedStrings_};
}

std::pair<bool, QVector<ColumnType>> ImportXlsx::getColumnTypes(
    const QString& sheetName)
{
    if (const auto it{columnTypes_.constFind(sheetName)};
        it != columnTypes_.constEnd())
        return {true, *it};

    if (!isCommonDataOk())
        return {false, {}};

    if (!isSheetNameValid(getSheetNames().second, sheetName))
        return {false, {}};

    if (!analyzeSheet(sheetName))
        return {false, {}};

    return {true, columnTypes_.value(sheetName)};
}

bool ImportXlsx::moveToSecondRow(QuaZipFile& zipFile,
                                 QXmlStreamReader& xmlStreamReader) const
{
    xmlStreamReader.setDevice(&zipFile);
    bool secondRow{false};
    while (!xmlStreamReader.atEnd())
    {
        if (isRowStart(xmlStreamReader))
        {
            if (secondRow)
                break;
            secondRow = true;
        }
        xmlStreamReader.readNextStartElement();
    }
    return true;
}

std::pair<bool, unsigned int> ImportXlsx::getCount(
    const QString& sheetName, const QHash<QString, unsigned int>& countMap)
{
    if (const auto it{countMap.find(sheetName)}; it != countMap.end())
        return {true, it.value()};

    if (!isCommonDataOk())
        return {false, {}};

    if (!isSheetNameValid(getSheetNames().second, sheetName))
        return {false, {}};

    if (!analyzeSheet(sheetName))
        return {false, 0};

    return {true, countMap.find(sheetName).value()};
}

std::pair<bool, QDomNodeList> ImportXlsx::getSheetNodes(
    QuaZipFile& zipFile,
    const std::function<QDomNodeList(const QDomElement&)>& nodesRetriever)
{
    QDomDocument xmlDocument;
    if (!xmlDocument.setContent(zipFile.readAll()))
    {
        setError(QStringLiteral("File is corrupted."));
        return {false, {}};
    }

    const QDomElement root{xmlDocument.documentElement()};
    const QDomNodeList sheetNodes{nodesRetriever(root)};
    if (sheetNodes.size() <= 0)
    {
        setError(QStringLiteral("File is corrupted, no sheets in xml."));
        return {false, {}};
    }
    return {true, sheetNodes};
}

bool ImportXlsx::isRowStart(const QXmlStreamReader& xmlStreamReader) const
{
    return (xmlStreamReader.name() == ROW_TAG) &&
           xmlStreamReader.isStartElement();
}

bool ImportXlsx::isCellStart(const QXmlStreamReader& xmlStreamReader) const
{
    return (xmlStreamReader.name() == CELL_TAG) &&
           xmlStreamReader.isStartElement();
}

bool ImportXlsx::isCellEnd(const QXmlStreamReader& xmlStreamReader) const
{
    return (xmlStreamReader.name() == CELL_TAG) &&
           xmlStreamReader.isEndElement();
}

bool ImportXlsx::isVTagStart(const QXmlStreamReader& xmlStreamReader) const
{
    return (xmlStreamReader.name() == V_TAG) &&
           xmlStreamReader.isStartElement();
}

std::pair<bool, QStringList> ImportXlsx::retrieveColumnNames(
    const QString& sheetName)
{
    auto [sheetFound, sheetPath]{getSheetPath(sheetName)};
    if (!sheetFound)
        return {false, {}};

    QuaZipFile zipFile;
    if (!initZipFile(zipFile, sheetPath))
        return {false, {}};

    QXmlStreamReader xmlStreamReader;
    xmlStreamReader.setDevice(&zipFile);
    skipToFirstRow(xmlStreamReader);

    int columnIndex{NOT_SET_COLUMN};
    QXmlStreamReader::TokenType lastToken{xmlStreamReader.tokenType()};
    QStringList columnNames;
    QString currentColType;
    while ((!xmlStreamReader.atEnd()) && (xmlStreamReader.name() != ROW_TAG))
    {
        if (isCellStart(xmlStreamReader))
        {
            ++columnIndex;
            const int currentColumnNumber{
                getCurrentColumnNumber(xmlStreamReader, 1)};
            while (currentColumnNumber > columnIndex)
            {
                columnNames << emptyColName_;
                ++columnIndex;
            }
            currentColType =
                xmlStreamReader.attributes().value(T_TAG).toString();
        }

        if ((!xmlStreamReader.atEnd()) && isVTagStart(xmlStreamReader))
            columnNames.append(getColumnName(xmlStreamReader, currentColType));

        // If we encounter empty cell than add it to list.
        if (isCellEnd(xmlStreamReader) &&
            (lastToken == QXmlStreamReader::StartElement))
            columnNames << emptyColName_;
        lastToken = xmlStreamReader.tokenType();
        xmlStreamReader.readNext();
    }
    return {true, columnNames};
}

std::tuple<bool, unsigned int, QVector<ColumnType>>
ImportXlsx::retrieveRowCountAndColumnTypes(const QString& sheetName)
{
    auto [sheetFound, sheetPath] = getSheetPath(sheetName);
    if (!sheetFound)
        return {false, {}, {}};

    QuaZipFile zipFile;
    if (!initZipFile(zipFile, sheetPath))
        return {false, {}, {}};

    QXmlStreamReader xmlStreamReader;
    if (!moveToSecondRow(zipFile, xmlStreamReader))
        return {false, {}, {}};

    QVector<ColumnType> columnTypes;
    int column{NOT_SET_COLUMN};
    int maxColumnIndex{NOT_SET_COLUMN};
    int rowCounter{0};
    int rowCountDigitsInXlsx{0};
    while ((!xmlStreamReader.atEnd()) &&
           (xmlStreamReader.name() != SHEET_DATA_TAG))
    {
        if (isRowStart(xmlStreamReader))
        {
            rowCountDigitsInXlsx = static_cast<int>(
                xmlStreamReader.attributes().value(R_TAG).size());
            ++rowCounter;
        }

        if (isCellStart(xmlStreamReader))
        {
            column =
                getCurrentColumnNumber(xmlStreamReader, rowCountDigitsInXlsx);
            maxColumnIndex = std::max(maxColumnIndex, column);
            if (column >= columnTypes.size())
                for (int i{static_cast<int>(columnTypes.size())}; i <= column;
                     ++i)
                    columnTypes.push_back(ColumnType::UNKNOWN);
            columnTypes[column] =
                recognizeColumnType(columnTypes.at(column), xmlStreamReader);
        }
        xmlStreamReader.readNextStartElement();
    }

    std::replace_if(
        columnTypes.begin(), columnTypes.end(), [](ColumnType columnType)
        { return columnType == ColumnType::UNKNOWN; }, ColumnType::STRING);

    return {true, rowCounter, columnTypes};
}

QString ImportXlsx::getColumnName(QXmlStreamReader& xmlStreamReader,
                                  const QString& currentColType) const
{
    QString columnName;
    if (currentColType == S_TAG)
    {
        const int value{xmlStreamReader.readElementText().toInt()};
        columnName = (*sharedStrings_)[value];
    }
    else
    {
        columnName = xmlStreamReader.readElementText();
    }
    return columnName;
}

void ImportXlsx::skipToFirstRow(QXmlStreamReader& xmlStreamReader) const
{
    while ((!xmlStreamReader.atEnd()) &&
           (xmlStreamReader.name() != SHEET_DATA_TAG))
        xmlStreamReader.readNext();
    xmlStreamReader.readNext();
    xmlStreamReader.readNext();
}

int ImportXlsx::getCurrentColumnNumber(const QXmlStreamReader& xmlStreamReader,
                                       int rowCountDigitsInXlsx) const
{
    QString rowNumber{xmlStreamReader.attributes().value(R_TAG).toString()};
    rowNumber.chop(rowCountDigitsInXlsx);
    return excelColNames_[rowNumber.toUtf8()];
}

ColumnType ImportXlsx::recognizeColumnType(
    ColumnType currentType, const QXmlStreamReader& xmlStreamReader) const
{
    if (ColumnType::STRING == currentType)
        return currentType;

    const QXmlStreamAttributes attributes{xmlStreamReader.attributes()};
    if (const QString value{attributes.value(T_TAG).toString()};
        (value == S_TAG) || (value == STR_TAG))
        return ColumnType::STRING;

    ColumnType detectedType{currentType};
    if (const QString sTagValue{attributes.value(S_TAG).toString()};
        isDateStyle(sTagValue))
    {
        if (currentType == ColumnType::UNKNOWN)
            detectedType = ColumnType::DATE;
        else
        {
            if (currentType != ColumnType::DATE)
                detectedType = ColumnType::STRING;
        }
    }
    else
    {
        if (currentType == ColumnType::UNKNOWN)
            detectedType = ColumnType::NUMBER;
    }
    return detectedType;
}

std::pair<bool, QMap<QString, QString>>
ImportXlsx::getSheetIdToUserFriendlyNameMap()
{
    QuaZipFile zipFile;
    if (!initZipFile(zipFile, QStringLiteral("xl/workbook.xml")))
        return {false, {}};

    auto nodesRetriever{[](const QDomElement& root) {
        return root.firstChildElement(QStringLiteral("sheets")).childNodes();
    }};
    auto [sucess, sheetNodes]{getSheetNodes(zipFile, nodesRetriever)};
    if (!sucess)
        return {false, {}};

    QMap<QString, QString> sheetIdToUserFriendlyNameMap;
    const int sheetNodesCount{sheetNodes.size()};
    for (int i{0}; i < sheetNodesCount; ++i)
    {
        const QDomElement sheet{sheetNodes.at(i).toElement()};
        if (!sheet.isNull())
            sheetIdToUserFriendlyNameMap[sheet.attribute(QStringLiteral(
                "r:id"))] = sheet.attribute(QStringLiteral("name"));
    }
    return {true, sheetIdToUserFriendlyNameMap};
}

std::pair<bool, QVector<std::pair<QString, QString>>>
ImportXlsx::retrieveSheets(
    const QMap<QString, QString>& sheetIdToUserFriendlyNameMap)
{
    QuaZipFile zipFile;
    if (!initZipFile(zipFile, QStringLiteral("xl/_rels/workbook.xml.rels")))
        return {false, {}};

    auto [sucess, sheetNodes]{getSheetNodes(
        zipFile, [](const QDomElement& root) { return root.childNodes(); })};
    if (!sucess)
        return {false, {}};

    QVector<std::pair<QString, QString>> sheets;
    const int sheetNodesCount{sheetNodes.size()};
    for (int i{0}; i < sheetNodesCount; ++i)
    {
        const QDomElement sheet{sheetNodes.at(i).toElement()};
        if (!sheet.isNull())
        {
            const auto it{sheetIdToUserFriendlyNameMap.constFind(
                sheet.attribute(QStringLiteral("Id")))};
            if (sheetIdToUserFriendlyNameMap.constEnd() != it)
                sheets.append(
                    {*it, "xl/" + sheet.attribute(QStringLiteral("Target"))});
        }
    }
    return {true, sheets};
}

std::pair<bool, QList<int>> ImportXlsx::getAllStyles()
{
    if (allStyles_)
        return {true, *allStyles_};
    bool success{false};
    std::tie(success, dateStyles_, allStyles_) = getStyles();
    return {success, *allStyles_};
}

QVariant ImportXlsx::getCurrentValue(QXmlStreamReader& xmlStreamReader,
                                     ColumnType currentColumnFormat,
                                     const QString& currentColType,
                                     const QString& actualSTagValue) const
{
    const QString valueAsString{xmlStreamReader.readElementText()};
    QVariant currentValue;
    switch (currentColumnFormat)
    {
        case ColumnType::STRING:
        {
            if (currentColType == S_TAG)
                currentValue = valueAsString.toInt();
            else
            {
                if (isDateStyle(actualSTagValue))
                    currentValue =
                        getDateFromString(valueAsString).toString(Qt::ISODate);
                else
                    currentValue = valueAsString;
            }
            break;
        }

        case ColumnType::DATE:
        {
            currentValue = getDateFromString(valueAsString);
            break;
        }

        case ColumnType::NUMBER:
        {
            currentValue = valueAsString.toDouble();
            break;
        }

        default:
        {
            Q_ASSERT(false);
            break;
        }
    }
    return currentValue;
}

QDate ImportXlsx::getDateFromString(const QString& dateAsString)
{
    const int daysToAdd{static_cast<int>(dateAsString.toDouble())};
    return utilities::getStartOfExcelWorld().addDays(daysToAdd);
}

bool ImportXlsx::isDateStyle(const QString& sTagValue) const
{
    return (!sTagValue.isEmpty()) &&
           dateStyles_->contains(allStyles_->at(sTagValue.toInt()));
}

bool ImportXlsx::isCommonDataOk()
{
    if ((!sheets_) && (!getSheetNames().first))
        return false;

    if ((!sharedStrings_) && (!getSharedStrings().first))
        return false;

    if ((!getDateStyles().first) && (!getAllStyles().first))
        return false;

    return true;
}

std::pair<bool, QVector<QVector<QVariant>>> ImportXlsx::getLimitedData(
    const QString& sheetName, const QVector<unsigned int>& excludedColumns,
    unsigned int rowLimit)
{
    const auto [success, columnTypes]{getColumnTypes(sheetName)};
    if (!success)
        return {false, {}};

    const unsigned int columnCount{columnCounts_.value(sheetName)};
    if (!columnsToExcludeAreValid(excludedColumns, columnCount))
        return {false, {}};

    auto [sheetFound, sheetPath]{getSheetPath(sheetName)};
    if (!sheetFound)
        return {false, {}};

    QuaZipFile zipFile;
    if (!initZipFile(zipFile, sheetPath))
        return {false, {}};

    QXmlStreamReader xmlStreamReader;
    moveToSecondRow(zipFile, xmlStreamReader);

    const QVector<QVariant> templateDataRow(
        createTemplateDataRow(excludedColumns, columnTypes));

    QMap<unsigned int, unsigned int> activeColumnsMapping{
        createActiveColumnMapping(excludedColumns, columnCount)};

    QVector<QVector<QVariant>> dataContainer(
        static_cast<int>(std::min(getRowCount(sheetName).second, rowLimit)));

    QVector<QVariant> currentDataRow(templateDataRow);
    int column{NOT_SET_COLUMN};
    QString currentColType{S_TAG};
    QString actualSTagValue;
    unsigned int rowCounter{0};
    unsigned int lastEmittedPercent{0};
    int rowCountDigitsInXlsx{0};
    while ((!xmlStreamReader.atEnd()) &&
           (xmlStreamReader.name() != SHEET_DATA_TAG) &&
           (rowCounter <= rowLimit))
    {
        if (isRowStart(xmlStreamReader))
        {
            column = NOT_SET_COLUMN;
            if (rowCounter != 0)
            {
                dataContainer[static_cast<int>(rowCounter - 1)] =
                    currentDataRow;
                currentDataRow = QVector<QVariant>(templateDataRow);
            }
            rowCountDigitsInXlsx = static_cast<int>(
                xmlStreamReader.attributes().value(R_TAG).size());
            ++rowCounter;
            updateProgress(rowCounter, rowLimit, lastEmittedPercent);
        }

        if (isCellStart(xmlStreamReader))
        {
            column =
                getCurrentColumnNumber(xmlStreamReader, rowCountDigitsInXlsx);
            currentColType =
                xmlStreamReader.attributes().value(T_TAG).toString();
            actualSTagValue =
                xmlStreamReader.attributes().value(S_TAG).toString();
        }

        if ((!xmlStreamReader.atEnd()) && isVTagStart(xmlStreamReader) &&
            (!excludedColumns.contains(static_cast<unsigned int>(column))))
        {
            const ColumnType format{columnTypes.at(column)};
            const unsigned int mappedColumnIndex{
                activeColumnsMapping[static_cast<unsigned int>(column)]};
            currentDataRow[static_cast<int>(mappedColumnIndex)] =
                getCurrentValue(xmlStreamReader, format, currentColType,
                                actualSTagValue);
        }
        xmlStreamReader.readNextStartElement();
    }

    if (currentDataRow != templateDataRow)
        dataContainer[dataContainer.size() - 1] = currentDataRow;

    return {true, dataContainer};
}

std::pair<bool, QList<int>> ImportXlsx::getDateStyles()
{
    if (dateStyles_)
        return {true, *dateStyles_};
    bool success{false};
    std::tie(success, dateStyles_, allStyles_) = getStyles();
    return {success, *dateStyles_};
}
