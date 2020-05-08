#include "ImportXlsx.h"

#include <algorithm>
#include <cmath>

#include <quazip5/quazip.h>
#include <quazip5/quazipfile.h>
#include <QIODevice>
#include <QMap>
#include <QSet>
#include <QVariant>
#include <QXmlStreamReader>
#include <QtXml/QDomDocument>

#include "EibleUtilities.h"

const QString ImportXlsx::ROW_TAG{QStringLiteral("row")};
const QString ImportXlsx::CELL_TAG{QStringLiteral("c")};
const QString ImportXlsx::SHEET_DATA_TAG{QStringLiteral("sheetData")};
const QString ImportXlsx::S_TAG{QStringLiteral("s")};
const QString ImportXlsx::V_TAG{QStringLiteral("v")};
const QString ImportXlsx::R_TAG{QStringLiteral("r")};
const QString ImportXlsx::T_TAG{QStringLiteral("t")};
const QString ImportXlsx::STR_TAG{QStringLiteral("str")};

ImportXlsx::ImportXlsx(QIODevice& ioDevice)
    : ImportSpreadsheet(ioDevice),
      excelColNames_(EibleUtilities::generateExcelColumnNames(
          EibleUtilities::getMaxExcelColumns()))
{
}

std::pair<bool, QStringList> ImportXlsx::getSheetNames()
{
    if (sheets_)
    {
        QStringList sheets;
        for (const auto& [sheetName, sheetPath] : *sheets_)
            sheets << sheetName;
        return {true, sheets};
    }

    auto [success, sheetIdToUserFriendlyNameMap] =
        getSheetIdToUserFriendlyNameMap();
    if (!success)
        return {false, {}};

    std::tie(success, sheets_) = retrieveSheets(sheetIdToUserFriendlyNameMap);
    if (!success)
        return {false, {}};

    QStringList sheetsToReturn;
    for (const auto& [sheetName, sheetPath] : *sheets_)
        sheetsToReturn << sheetName;
    return {true, sheetsToReturn};
}

std::pair<bool, QStringList> ImportXlsx::getColumnNames(
    const QString& sheetName)
{
    if (!sheets_ && !getSheetNames().first)
        return {false, {}};

    if (!sharedStrings_ && !getSharedStrings().first)
        return {false, {}};

    const auto it = columnCounts_.find(sheetName);
    if (it == columnCounts_.end() && !analyzeSheet(sheetName))
        return {false, {}};

    auto [sheetFound, sheetPath] = getSheetPath(sheetName);
    if (!sheetFound)
        return {false, {}};

    QuaZipFile zipFile;
    if (!initZipFile(zipFile, sheetPath))
        return {false, {}};

    return {true, retrieveColumnNames(zipFile, sheetName)};
}

QList<int> ImportXlsx::retrieveDateStyles(const QDomNodeList& sheetNodes) const
{
    QList<int> dateStyles;
    const QList predefinedExcelStylesForDates{14, 15, 16, 17, 22};
    dateStyles.append(predefinedExcelStylesForDates);
    for (int i = 0; i < sheetNodes.size(); ++i)
    {
        QDomElement sheet = sheetNodes.at(i).toElement();

        if (sheet.hasAttribute(QStringLiteral("numFmtId")) &&
            sheet.hasAttribute(QStringLiteral("formatCode")))
        {
            QString formatCode = sheet.attribute(QStringLiteral("formatCode"));
            bool gotD =
                formatCode.contains(QStringLiteral("d"), Qt::CaseInsensitive);
            bool gotM =
                formatCode.contains(QStringLiteral("m"), Qt::CaseInsensitive);
            bool gotY =
                formatCode.contains(QStringLiteral("y"), Qt::CaseInsensitive);

            if ((gotD && gotY) || (gotD && gotM) || (gotM && gotY))
                dateStyles.push_back(
                    sheet.attribute(QStringLiteral("numFmtId")).toInt());
        }
    }
    return dateStyles;
}

QList<int> ImportXlsx::retrieveAllStyles(const QDomNodeList& sheetNodes) const
{
    QList<int> allStyles;
    for (int i = 0; i < sheetNodes.size(); ++i)
    {
        QDomElement sheet = sheetNodes.at(i).toElement();

        if (!sheet.isNull() && sheet.hasAttribute(QStringLiteral("numFmtId")))
            allStyles.push_back(
                sheet.attribute(QStringLiteral("numFmtId")).toInt());
    }
    return allStyles;
}

std::tuple<bool, std::optional<QList<int>>, std::optional<QList<int>>>
ImportXlsx::getStyles()
{
    QuaZipFile zipFile;
    if (!initZipFile(zipFile, QStringLiteral("xl/styles.xml")))
        return {false, {}, {}};

    QDomDocument xmlDocument(__FUNCTION__);
    if (!xmlDocument.setContent(zipFile.readAll()))
    {
        setError(__FUNCTION__, "Xml file is corrupted.");
        return {false, std::nullopt, std::nullopt};
    }

    QDomElement root = xmlDocument.documentElement();
    QDomNodeList sheetNodes =
        root.firstChildElement(QStringLiteral("numFmts")).childNodes();
    QList<int> dateStyles{retrieveDateStyles(sheetNodes)};

    sheetNodes = root.firstChildElement(QStringLiteral("cellXfs")).childNodes();
    QList<int> allStyles{retrieveAllStyles(sheetNodes)};

    return {true, dateStyles, allStyles};
}

std::pair<bool, QString> ImportXlsx::getSheetPath(QString sheetName)
{
    for (const auto& [currentSheetName, sheetPath] : *sheets_)
        if (currentSheetName == sheetName)
            return {true, sheetPath};

    QStringList sheetNames;
    for (const auto& [currenSheetName, sheetPath] : *sheets_)
        sheetNames << currenSheetName;
    setError(__FUNCTION__,
             "Can not find sheet path for sheet name " + sheetName +
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
    if (const auto it = columnTypes_.find(sheetName); it != columnTypes_.end())
        return {true, *it};

    if (!sheets_ && !getSheetNames().first)
        return {false, {}};

    if (!sharedStrings_ && !getSharedStrings().first)
        return {false, {}};

    if (!getDateStyles().first && !getAllStyles().first)
        return {false, {}};

    auto [sheetFound, sheetPath] = getSheetPath(sheetName);
    if (!sheetFound)
        return {false, {}};

    QuaZipFile zipFile;
    if (!initZipFile(zipFile, sheetPath))
        return {false, {}};

    QXmlStreamReader xmlStreamReader;
    if (!moveToSecondRow(zipFile, xmlStreamReader))
        return {false, {}};

    QVector<ColumnType> columnTypes;

    // Current column.
    int column = NOT_SET_COLUMN;
    int maxColumn = NOT_SET_COLUMN;

    // Current row.
    int rowCounter = 0;

    int charsToChop = 1;

    QXmlStreamAttributes xmlStreamAtrributes;

    while (!xmlStreamReader.atEnd() &&
           0 != xmlStreamReader.name().compare(SHEET_DATA_TAG))
    {
        // If start of row encountered than reset column counter add
        // increment row counter.
        if (isRowStart(xmlStreamReader))
        {
            column = NOT_SET_COLUMN;
            rowCounter++;
            double power = pow(DECIMAL_BASE, charsToChop);
            if (power <= rowCounter + 1)
                charsToChop++;
        }

        // When we encounter start of cell description.
        if (isCellStart(xmlStreamReader))
        {
            column++;
            column = getExpectedColumnIndex(xmlStreamReader, charsToChop);
            maxColumn = std::max(maxColumn, column);

            if (column >= columnTypes.size())
                for (int i = columnTypes.size(); i <= column; ++i)
                    columnTypes.push_back(ColumnType::UNKNOWN);

            // If data format in column is unknown than read it.
            if (columnTypes.at(column) == ColumnType::UNKNOWN)
            {
                xmlStreamAtrributes = xmlStreamReader.attributes();
                QString value = xmlStreamAtrributes.value(T_TAG).toString();
                bool valueIsSTag = (0 == value.compare(S_TAG));
                if (valueIsSTag || 0 == value.compare(STR_TAG))
                {
                    if (valueIsSTag)
                    {
                        xmlStreamReader.readNextStartElement();

                        if (isVTagStart(xmlStreamReader) &&
                            (*sharedStrings_)[xmlStreamReader.readElementText()
                                                  .toInt()]
                                .isEmpty())
                        {
                        }
                        else
                        {
                            columnTypes[column] = ColumnType::STRING;
                        }
                    }
                    else
                    {
                        columnTypes[column] = ColumnType::STRING;
                    }
                }
                else
                {
                    QString otherValue =
                        xmlStreamAtrributes.value(S_TAG).toString();
                    if (!otherValue.isEmpty() &&
                        dateStyles_->contains(
                            allStyles_->at(otherValue.toInt())))
                        columnTypes[column] = ColumnType::DATE;
                    else
                        columnTypes[column] = ColumnType::NUMBER;
                }
            }
            else
            {
                if (ColumnType::STRING != columnTypes.at(column))
                {
                    // If type of column is known than check if it is correct.
                    xmlStreamAtrributes = xmlStreamReader.attributes();
                    QString value = xmlStreamAtrributes.value(T_TAG).toString();
                    bool valueIsSTag = (0 == value.compare(S_TAG));
                    if (valueIsSTag || 0 == value.compare(STR_TAG))
                    {
                        if (valueIsSTag)
                        {
                            xmlStreamReader.readNextStartElement();

                            if (isVTagStart(xmlStreamReader) &&
                                (*sharedStrings_)
                                    [xmlStreamReader.readElementText().toInt()]
                                        .isEmpty())
                            {
                            }
                            else
                            {
                                columnTypes[column] = ColumnType::STRING;
                            }
                        }
                        else
                        {
                            columnTypes[column] = ColumnType::STRING;
                        }
                    }
                    else
                    {
                        QString othervalue =
                            xmlStreamAtrributes.value(S_TAG).toString();

                        if (!othervalue.isEmpty() &&
                            dateStyles_->contains(
                                allStyles_->at(othervalue.toInt())))
                        {
                            if (columnTypes.at(column) != ColumnType::DATE)
                                columnTypes[column] = ColumnType::STRING;
                        }
                        else
                        {
                            if (columnTypes.at(column) == ColumnType::UNKNOWN)
                                columnTypes[column] = ColumnType::STRING;
                        }
                    }
                }
            }
        }

        xmlStreamReader.readNextStartElement();
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

bool ImportXlsx::moveToSecondRow(QuaZipFile& zipFile,
                                 QXmlStreamReader& xmlStreamReader)
{
    xmlStreamReader.setDevice(&zipFile);

    bool secondRow = false;
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
    if (!sheets_ && !getSheetNames().first)
        return {false, {}};

    const auto it = countMap.find(sheetName);
    if (it != countMap.end())
        return {true, it.value()};

    if (!analyzeSheet(sheetName))
        return {false, 0};

    return {true, countMap.find(sheetName).value()};
}

int ImportXlsx::getExpectedColumnIndex(QXmlStreamReader& xmlStreamReader,
                                       unsigned int charsToChop) const
{
    QString stringToChop = xmlStreamReader.attributes().value(R_TAG).toString();
    int numberOfCharsToRemove = stringToChop.size() - charsToChop;
    int expectedColumnIndex = excelColNames_.indexOf(
        stringToChop.left(numberOfCharsToRemove).toUtf8());
    return expectedColumnIndex;
}

std::pair<unsigned int, unsigned int> ImportXlsx::getRowAndColumnCount(
    QXmlStreamReader& xmlStreamReader) const
{
    unsigned int rowCounter = 0;
    int maxColumnIndex = NOT_SET_COLUMN;
    int column = NOT_SET_COLUMN;
    unsigned int charsToChop = 1;

    while (!xmlStreamReader.atEnd() &&
           0 != xmlStreamReader.name().compare(SHEET_DATA_TAG))
    {
        if (isRowStart(xmlStreamReader))
        {
            rowCounter++;
            column = NOT_SET_COLUMN;
            double power = pow(DECIMAL_BASE, charsToChop);
            if (power <= rowCounter + 1)
                charsToChop++;
        }

        if (isCellStart(xmlStreamReader))
        {
            column = getExpectedColumnIndex(xmlStreamReader, charsToChop);
            maxColumnIndex = std::max(maxColumnIndex, column);
        }
        xmlStreamReader.readNextStartElement();
    }
    return {rowCounter, maxColumnIndex + 1};
}

bool ImportXlsx::analyzeSheet(const QString& sheetName)
{
    auto [sheetFound, sheetPath] = getSheetPath(sheetName);
    if (!sheetFound)
        return false;

    QuaZipFile zipFile;
    if (!initZipFile(zipFile, sheetPath))
        return false;

    QXmlStreamReader xmlStreamReader;
    if (!moveToSecondRow(zipFile, xmlStreamReader))
        return false;

    std::tie(rowCounts_[sheetName], columnCounts_[sheetName]) =
        getRowAndColumnCount(xmlStreamReader);

    return true;
}

std::pair<bool, QDomNodeList> ImportXlsx::getSheetNodes(
    QuaZipFile& zipFile,
    std::function<QDomNodeList(QDomElement)> nodesRetriever)
{
    QDomDocument xmlDocument = QDomDocument(__FUNCTION__);
    if (!xmlDocument.setContent(zipFile.readAll()))
    {
        setError(__FUNCTION__, "File is corrupted.");
        return {false, {}};
    }

    QDomElement root = xmlDocument.documentElement();
    QDomNodeList sheetNodes = nodesRetriever(root);

    if (sheetNodes.size() <= 0)
    {
        setError(__FUNCTION__, "File is corrupted, no sheets in xml.");
        return {false, {}};
    }
    return {true, sheetNodes};
}

bool ImportXlsx::isRowStart(const QXmlStreamReader& xmlStreamReader) const
{
    return 0 == xmlStreamReader.name().compare(ROW_TAG) &&
           xmlStreamReader.isStartElement();
}

bool ImportXlsx::isCellStart(const QXmlStreamReader& xmlStreamReader) const
{
    return xmlStreamReader.name().toString() == CELL_TAG &&
           xmlStreamReader.isStartElement();
}

bool ImportXlsx::isCellEnd(const QXmlStreamReader& xmlStreamReader) const
{
    return xmlStreamReader.name().toString() == CELL_TAG &&
           xmlStreamReader.isEndElement();
}

bool ImportXlsx::isVTagStart(const QXmlStreamReader& xmlStreamReader) const
{
    return 0 == xmlStreamReader.name().compare(V_TAG) &&
           xmlStreamReader.isStartElement();
}

QString ImportXlsx::getColumnName(QXmlStreamReader& xmlStreamReader,
                                  const QString& currentColType) const
{
    QString columnName;
    if (currentColType == S_TAG)
    {
        int value = xmlStreamReader.readElementText().toInt();
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
    while (!xmlStreamReader.atEnd() && xmlStreamReader.name() != SHEET_DATA_TAG)
        xmlStreamReader.readNext();
    xmlStreamReader.readNext();
    xmlStreamReader.readNext();
}

int ImportXlsx::getCurrentColumnNumber(
    const QXmlStreamReader& xmlStreamReader) const
{
    const QRegExp regExp(QLatin1String("[0-9]"));
    QString rowNumber(xmlStreamReader.attributes().value(R_TAG).toString());
    return excelColNames_.indexOf(rowNumber.remove(regExp).toUtf8());
}

QStringList ImportXlsx::retrieveColumnNames(QuaZipFile& zipFile,
                                            const QString& sheetName)
{
    QXmlStreamReader xmlStreamReader;
    xmlStreamReader.setDevice(&zipFile);
    skipToFirstRow(xmlStreamReader);

    int columnIndex = NOT_SET_COLUMN;
    QXmlStreamReader::TokenType lastToken = xmlStreamReader.tokenType();
    QStringList columnNames;
    QString currentColType;
    while (!xmlStreamReader.atEnd() && xmlStreamReader.name() != ROW_TAG)
    {
        if (isCellStart(xmlStreamReader))
        {
            columnIndex++;
            const int currentColumnNumber{
                getCurrentColumnNumber(xmlStreamReader)};
            while (currentColumnNumber > columnIndex)
            {
                columnNames << emptyColName_;
                columnIndex++;
            }
            currentColType =
                xmlStreamReader.attributes().value(T_TAG).toString();
        }

        if (!xmlStreamReader.atEnd() && isVTagStart(xmlStreamReader))
            columnNames.append(getColumnName(xmlStreamReader, currentColType));

        // If we encounter empty cell than add it to list.
        if (isCellEnd(xmlStreamReader) &&
            lastToken == QXmlStreamReader::StartElement)
            columnNames << emptyColName_;
        lastToken = xmlStreamReader.tokenType();
        xmlStreamReader.readNext();
    }

    const unsigned int columnCount{columnCounts_.find(sheetName).value()};
    while (columnNames.count() < static_cast<int>(columnCount))
        columnNames << emptyColName_;

    return columnNames;
}

std::pair<bool, QMap<QString, QString>>
ImportXlsx::getSheetIdToUserFriendlyNameMap()
{
    QuaZipFile zipFile;
    if (!initZipFile(zipFile, QStringLiteral("xl/workbook.xml")))
        return {false, {}};

    auto nodesRetriever{[](QDomElement root) {
        return root.firstChildElement(QStringLiteral("sheets")).childNodes();
    }};
    auto [sucess, sheetNodes] = getSheetNodes(zipFile, nodesRetriever);
    if (!sucess)
        return {false, {}};

    QMap<QString, QString> sheetIdToUserFriendlyNameMap;
    for (int i = 0; i < sheetNodes.size(); ++i)
    {
        QDomElement sheet = sheetNodes.at(i).toElement();
        if (!sheet.isNull())
            sheetIdToUserFriendlyNameMap[sheet.attribute(QStringLiteral(
                "r:id"))] = sheet.attribute(QStringLiteral("name"));
    }
    return {true, sheetIdToUserFriendlyNameMap};
}

std::pair<bool, QList<std::pair<QString, QString>>> ImportXlsx::retrieveSheets(
    const QMap<QString, QString>& sheetIdToUserFriendlyNameMap)
{
    QuaZipFile zipFile;
    if (!initZipFile(zipFile, QStringLiteral("xl/_rels/workbook.xml.rels")))
        return {false, {}};

    auto [sucess, sheetNodes] = getSheetNodes(
        zipFile, [](QDomElement root) { return root.childNodes(); });
    if (!sucess)
        return {false, {}};

    QList<std::pair<QString, QString>> sheets;
    for (int i = 0; i < sheetNodes.size(); ++i)
    {
        const QDomElement sheet = sheetNodes.at(i).toElement();

        if (!sheet.isNull())
        {
            const auto it = sheetIdToUserFriendlyNameMap.constFind(
                sheet.attribute(QStringLiteral("Id")));
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

std::pair<bool, unsigned int> ImportXlsx::getColumnCount(
    const QString& sheetName)
{
    return getCount(sheetName, columnCounts_);
}

std::pair<bool, unsigned int> ImportXlsx::getRowCount(const QString& sheetName)
{
    return getCount(sheetName, rowCounts_);
}

std::pair<bool, QVector<QVector<QVariant>>> ImportXlsx::getLimitedData(
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

    auto [sheetFound, sheetPath] = getSheetPath(sheetName);
    if (!sheetFound)
        return {false, {}};

    QuaZipFile zipFile;
    if (!initZipFile(zipFile, sheetPath))
        return {false, {}};

    QXmlStreamReader xmlStreamReader;
    moveToSecondRow(zipFile, xmlStreamReader);

    // Actual column number.
    int column = NOT_SET_COLUMN;

    // Actual data type in cell (s, str, null).
    QString currentColType = S_TAG;
    QString actualSTagValue{};

    // Actual row number.
    unsigned int rowCounter = 0;

    // Null elements row.
    QVector<QVariant> templateDataRow;

    QMap<int, int> activeColumnsMapping = QMap<int, int>();

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

    // Number of chars to chop from end.
    int charsToChop = 1;

    unsigned int lastEmittedPercent{0};
    while (!xmlStreamReader.atEnd() &&
           0 != xmlStreamReader.name().compare(SHEET_DATA_TAG) &&
           rowCounter <= rowLimit)
    {
        // If start of row encountered than reset column counter add
        // increment row counter.
        if (isRowStart(xmlStreamReader))
        {
            column = NOT_SET_COLUMN;

            if (0 != rowCounter)
            {
                dataContainer[rowCounter - 1] = currentDataRow;
                currentDataRow = QVector<QVariant>(templateDataRow);

                double power = pow(DECIMAL_BASE, charsToChop);
                if (static_cast<unsigned int>(qRound(power)) <= rowCounter + 2)
                    charsToChop++;
            }

            rowCounter++;

            updateProgress(rowCounter, rowLimit, lastEmittedPercent);

            if (fillSamplesOnly && rowCounter > containerSize)
                break;
        }

        // When we encounter start of cell description.
        if (isCellStart(xmlStreamReader))
        {
            column++;
            column = getExpectedColumnIndex(xmlStreamReader, charsToChop);

            // Remember cell type.
            currentColType =
                xmlStreamReader.attributes().value(T_TAG).toString();
            actualSTagValue =
                xmlStreamReader.attributes().value(S_TAG).toString();
        }

        // Insert data into table.
        if (!xmlStreamReader.atEnd() && isVTagStart(xmlStreamReader) &&
            (fillSamplesOnly || !excludedColumns.contains(column)))
        {
            ColumnType format = columnTypes.at(column);

            switch (format)
            {
                case ColumnType::STRING:
                {
                    // Strings.
                    if (0 == currentColType.compare(S_TAG))
                        currentDataRow[activeColumnsMapping[column]] =
                            QVariant(xmlStreamReader.readElementText().toInt());
                    else
                    {
                        if (!actualSTagValue.isEmpty() &&
                            dateStyles_->contains(
                                allStyles_->at(actualSTagValue.toInt())))
                        {
                            int daysToAdd = static_cast<int>(
                                xmlStreamReader.readElementText().toInt());
                            currentDataRow[activeColumnsMapping[column]] =
                                QVariant(EibleUtilities::getStartOfExcelWorld()
                                             .addDays(daysToAdd)
                                             .toString(Qt::ISODate));
                        }
                        else
                        {
                            currentDataRow[activeColumnsMapping[column]] =
                                QVariant(xmlStreamReader.readElementText());
                        }
                    }
                    break;
                }

                case ColumnType::DATE:
                {
                    // If YYYY-MM-DD HH:MI:SS cut and left YYYY-MM-DD.
                    int daysToAdd = static_cast<int>(
                        xmlStreamReader.readElementText().toDouble());
                    currentDataRow[activeColumnsMapping[column]] =
                        QVariant(EibleUtilities::getStartOfExcelWorld().addDays(
                            daysToAdd));
                    break;
                }

                case ColumnType::NUMBER:
                {
                    if (0 != currentColType.compare(S_TAG))
                        currentDataRow[activeColumnsMapping[column]] = QVariant(
                            xmlStreamReader.readElementText().toDouble());
                    break;
                }

                case ColumnType::UNKNOWN:
                {
                    Q_ASSERT(false);
                    break;
                }
            }
        }
        xmlStreamReader.readNextStartElement();
    }

    if (rowCounter != 0 && (!fillSamplesOnly || rowLimit > rowCount))
    {
        Q_ASSERT(rowCounter <= rowCount);
        if (rowCounter <= rowCount)
            dataContainer[rowCounter - 1] = currentDataRow;
    }

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
