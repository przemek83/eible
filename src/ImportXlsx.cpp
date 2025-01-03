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
    if (!sheets_)
    {
        auto [success,
              sheetIdToUserFriendlyNameMap]{getSheetIdToUserFriendlyNameMap()};
        if (!success)
            return {false, {}};

        std::tie(success, sheets_) =
            retrieveSheets(sheetIdToUserFriendlyNameMap);
        if (!success)
            return {false, {}};
    }

    return {true, createSheetNames()};
}

std::pair<bool, QStringList> ImportXlsx::getColumnNames(
    const QString& sheetName)
{
    if (const auto it{columnNames_.constFind(sheetName)};
        it != columnNames_.constEnd())
        return {true, *it};

    if (isCommonDataOk() &&
        isSheetNameValid(getSheetNames().second, sheetName) &&
        analyzeSheet(sheetName))
        return {true, columnNames_.value(sheetName)};

    return {false, {}};
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
    QuaZipFile quaZipFile;
    if (!initZipFile(quaZipFile, QStringLiteral("xl/styles.xml")))
        return {false, {}, {}};

    QDomDocument xmlDocument;
    if (!xmlDocument.setContent(quaZipFile.readAll()))
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

    QStringList sheetNames{createSheetNames()};
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

    QuaZipFile quaZipFile;
    if (!openZipFile(quaZipFile))
        return {false, {}};

    QXmlStreamReader reader;
    reader.setDevice(&quaZipFile);

    QStringList sharedStrings;
    while (!reader.atEnd())
    {
        reader.readNext();
        if (reader.name() == T_TAG)
            sharedStrings.append(reader.readElementText());
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

    if (isCommonDataOk() &&
        isSheetNameValid(getSheetNames().second, sheetName) &&
        analyzeSheet(sheetName))
        return {true, columnTypes_.value(sheetName)};

    return {false, {}};
}

void ImportXlsx::moveToSecondRow(QuaZipFile& quaZipFile,
                                 QXmlStreamReader& reader) const
{
    reader.setDevice(&quaZipFile);
    bool secondRow{false};
    while (!reader.atEnd())
    {
        if (isRowStart(reader))
        {
            if (secondRow)
                return;

            secondRow = true;
        }
        reader.readNextStartElement();
    }
}

std::pair<bool, int> ImportXlsx::getCount(const QString& sheetName,
                                          const QHash<QString, int>& countMap)
{
    if (const auto it{countMap.find(sheetName)}; it != countMap.end())
        return {true, *it};

    if (isCommonDataOk() &&
        isSheetNameValid(getSheetNames().second, sheetName) &&
        analyzeSheet(sheetName))
        return {true, countMap.find(sheetName).value()};

    return {false, {}};
}

std::pair<bool, QDomNodeList> ImportXlsx::getSheetNodes(
    QuaZipFile& quaZipFile,
    const std::function<QDomNodeList(const QDomElement&)>& nodesRetriever)
{
    QDomDocument xmlDocument;
    if (!xmlDocument.setContent(quaZipFile.readAll()))
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

bool ImportXlsx::isRowStart(const QXmlStreamReader& reader) const
{
    return (reader.name() == ROW_TAG) && reader.isStartElement();
}

bool ImportXlsx::isCellStart(const QXmlStreamReader& reader) const
{
    return (reader.name() == CELL_TAG) && reader.isStartElement();
}

bool ImportXlsx::isCellEnd(const QXmlStreamReader& reader) const
{
    return (reader.name() == CELL_TAG) && reader.isEndElement();
}

bool ImportXlsx::isVTagStart(const QXmlStreamReader& reader) const
{
    return (reader.name() == V_TAG) && reader.isStartElement();
}

std::pair<bool, QStringList> ImportXlsx::retrieveColumnNames(
    const QString& sheetName)
{
    auto [sheetFound, sheetPath]{getSheetPath(sheetName)};
    if (!sheetFound)
        return {false, {}};

    QuaZipFile quaZipFile;
    if (!initZipFile(quaZipFile, sheetPath))
        return {false, {}};

    QXmlStreamReader reader;
    reader.setDevice(&quaZipFile);
    skipToFirstRow(reader);

    int columnIndex{NOT_SET_COLUMN};
    QXmlStreamReader::TokenType lastToken{reader.tokenType()};
    QStringList columnNames;
    QString currentColType;
    while ((!reader.atEnd()) && (reader.name() != ROW_TAG))
    {
        if (isCellStart(reader))
        {
            ++columnIndex;
            const int currentColumnNumber{getCurrentColumnNumber(reader, 1)};
            while (currentColumnNumber > columnIndex)
            {
                columnNames << emptyColName_;
                ++columnIndex;
            }
            currentColType = reader.attributes().value(T_TAG).toString();
        }

        if ((!reader.atEnd()) && isVTagStart(reader))
            columnNames.append(getColumnName(reader, currentColType));

        // If we encounter empty cell than add it to list.
        if (isCellEnd(reader) && (lastToken == QXmlStreamReader::StartElement))
            columnNames << emptyColName_;
        lastToken = reader.tokenType();
        reader.readNext();
    }
    return {true, columnNames};
}

std::tuple<bool, int, QVector<ColumnType>>
ImportXlsx::retrieveRowCountAndColumnTypes(const QString& sheetName)
{
    auto [sheetFound, sheetPath] = getSheetPath(sheetName);
    if (!sheetFound)
        return {false, {}, {}};

    QuaZipFile quaZipFile;
    if (!initZipFile(quaZipFile, sheetPath))
        return {false, {}, {}};

    QXmlStreamReader reader;
    moveToSecondRow(quaZipFile, reader);

    QVector<ColumnType> columnTypes;
    int column{NOT_SET_COLUMN};
    int maxColumnIndex{NOT_SET_COLUMN};
    int rowCounter{0};
    int rowCountDigitsInXlsx{0};
    while ((!reader.atEnd()) && (reader.name() != SHEET_DATA_TAG))
    {
        if (isRowStart(reader))
        {
            rowCountDigitsInXlsx =
                static_cast<int>(reader.attributes().value(R_TAG).size());
            ++rowCounter;
        }

        if (isCellStart(reader))
        {
            column = getCurrentColumnNumber(reader, rowCountDigitsInXlsx);
            maxColumnIndex = std::max(maxColumnIndex, column);
            adjustColumnTypeSize(columnTypes, column);
            columnTypes[column] =
                recognizeColumnType(columnTypes.at(column), reader);
        }
        reader.readNextStartElement();
    }

    std::replace_if(
        columnTypes.begin(), columnTypes.end(), [](ColumnType columnType)
        { return columnType == ColumnType::UNKNOWN; }, ColumnType::STRING);

    return {true, rowCounter, columnTypes};
}

QString ImportXlsx::getColumnName(QXmlStreamReader& reader,
                                  const QString& currentColType) const
{
    if (currentColType != S_TAG)
        return reader.readElementText();

    const int value{reader.readElementText().toInt()};
    return sharedStrings_.value()[value];
}

void ImportXlsx::skipToFirstRow(QXmlStreamReader& reader) const
{
    while ((!reader.atEnd()) && (reader.name() != SHEET_DATA_TAG))
        reader.readNext();
    reader.readNext();
    reader.readNext();
}

int ImportXlsx::getCurrentColumnNumber(const QXmlStreamReader& reader,
                                       int rowCountDigitsInXlsx) const
{
    QString rowNumber{reader.attributes().value(R_TAG).toString()};
    rowNumber.chop(rowCountDigitsInXlsx);
    return excelColNames_[rowNumber.toUtf8()];
}

ColumnType ImportXlsx::recognizeColumnType(ColumnType currentType,
                                           const QXmlStreamReader& reader) const
{
    if (ColumnType::STRING == currentType)
        return currentType;

    const QXmlStreamAttributes attributes{reader.attributes()};
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
    QuaZipFile quaZipFile;
    if (!initZipFile(quaZipFile, QStringLiteral("xl/workbook.xml")))
        return {false, {}};

    auto nodesRetriever{[](const QDomElement& root) {
        return root.firstChildElement(QStringLiteral("sheets")).childNodes();
    }};
    auto [sucess, sheetNodes]{getSheetNodes(quaZipFile, nodesRetriever)};
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
    QuaZipFile quaZipFile;
    if (!initZipFile(quaZipFile, QStringLiteral("xl/_rels/workbook.xml.rels")))
        return {false, {}};

    auto [sucess, sheetNodes]{getSheetNodes(
        quaZipFile, [](const QDomElement& root) { return root.childNodes(); })};
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

QVariant ImportXlsx::getCurrentValueForStringColumn(
    const QString& currentColType, const QString& actualSTagValue,
    const QString& valueAsString) const
{
    if (currentColType == S_TAG)
        return valueAsString.toInt();

    if (isDateStyle(actualSTagValue))
        return getDateFromString(valueAsString).toString(Qt::ISODate);

    return valueAsString;
}

QVariant ImportXlsx::getCurrentValue(QXmlStreamReader& reader,
                                     ColumnType currentColumnFormat,
                                     const QString& currentColType,
                                     const QString& actualSTagValue) const
{
    const QString valueAsString{reader.readElementText()};
    QVariant currentValue;
    switch (currentColumnFormat)
    {
        case ColumnType::STRING:
        {
            currentValue = getCurrentValueForStringColumn(
                currentColType, actualSTagValue, valueAsString);
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

    return getDateStyles().first && getAllStyles().first;
}

std::pair<bool, QVector<QVector<QVariant>>> ImportXlsx::getLimitedData(
    const QString& sheetName, const QVector<int>& excludedColumns, int rowLimit)
{
    const auto [success, columnTypes]{getColumnTypes(sheetName)};
    if (!success)
        return {false, {}};

    const int columnCount{columnCounts_.value(sheetName)};
    if (!columnsToExcludeAreValid(excludedColumns, columnCount))
        return {false, {}};

    auto [sheetFound, sheetPath]{getSheetPath(sheetName)};
    if (!sheetFound)
        return {false, {}};

    QuaZipFile quaZipFile;
    if (!initZipFile(quaZipFile, sheetPath))
        return {false, {}};

    QXmlStreamReader reader;
    moveToSecondRow(quaZipFile, reader);

    const QVector<QVariant> templateDataRow(
        createTemplateDataRow(excludedColumns, columnTypes));

    QMap<int, int> activeColumnsMapping{
        createActiveColumnMapping(excludedColumns, columnCount)};

    QVector<QVector<QVariant>> dataContainer(
        std::min(getRowCount(sheetName).second, rowLimit));

    QVector<QVariant> currentDataRow(templateDataRow);
    int column{NOT_SET_COLUMN};
    QString currentColType{S_TAG};
    QString actualSTagValue;
    int rowCounter{0};
    int lastEmittedPercent{0};
    int rowCountDigitsInXlsx{0};
    while ((!reader.atEnd()) && (reader.name() != SHEET_DATA_TAG) &&
           (rowCounter <= rowLimit))
    {
        if (isRowStart(reader))
        {
            column = NOT_SET_COLUMN;
            if (rowCounter != 0)
            {
                dataContainer[rowCounter - 1] = currentDataRow;
                currentDataRow = QVector<QVariant>(templateDataRow);
            }
            rowCountDigitsInXlsx =
                static_cast<int>(reader.attributes().value(R_TAG).size());
            ++rowCounter;
            updateProgress(rowCounter, rowLimit, lastEmittedPercent);
        }

        if (isCellStart(reader))
        {
            column = getCurrentColumnNumber(reader, rowCountDigitsInXlsx);
            currentColType = reader.attributes().value(T_TAG).toString();
            actualSTagValue = reader.attributes().value(S_TAG).toString();
        }

        if ((!reader.atEnd()) && isVTagStart(reader) &&
            (!excludedColumns.contains(column)))
        {
            const ColumnType format{columnTypes.at(column)};
            const int mappedColumnIndex{activeColumnsMapping[column]};
            currentDataRow[mappedColumnIndex] = getCurrentValue(
                reader, format, currentColType, actualSTagValue);
        }
        reader.readNextStartElement();
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

QStringList ImportXlsx::createSheetNames() const
{
    QStringList sheetNames;
    sheetNames.reserve(sheets_->size());
    for (const auto& [sheetName, sheetPath] : *sheets_)
        sheetNames << sheetName;
    return sheetNames;
}

void ImportXlsx::adjustColumnTypeSize(QVector<ColumnType>& columnTypes,
                                      int column) const
{
    if (column >= columnTypes.size())
    {
        for (qsizetype i{columnTypes.size()}; i <= column; ++i)
            columnTypes.push_back(ColumnType::UNKNOWN);
    }
}
