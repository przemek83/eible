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

ImportXlsx::ImportXlsx(QIODevice& ioDevice) : ImportSpreadsheet(ioDevice) {}

std::pair<bool, QStringList> ImportXlsx::getSheetNames()
{
    if (sheets_)
    {
        QStringList sheets;
        for (const auto& [sheetName, sheetPath] : *sheets_)
            sheets << sheetName;
        return {true, sheets};
    }

    QuaZip zip(&ioDevice_);

    if (!zip.open(QuaZip::mdUnzip))
    {
        setError(__FUNCTION__,
                 "Can not open zip file " + zip.getZipName() + ".");
        return {false, {}};
    }

    QMap<QString, QString> sheetIdToUserFriendlyNameMap;
    QList<std::pair<QString, QString>> sheets;

    if (zip.setCurrentFile(QStringLiteral("xl/workbook.xml")))
    {
        // Open file in archive.
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
            setError(__FUNCTION__, "File is corrupted.");
            return {false, {}};
        }
        zipFile.close();

        QDomElement root = xmlDocument.documentElement();
        QDomNodeList sheetNodes =
            root.firstChildElement(QStringLiteral("sheets")).childNodes();

        if (sheetNodes.size() <= 0)
        {
            setError(__FUNCTION__, "File is corrupted, no sheets in xml.");
            return {false, {}};
        }

        for (int i = 0; i < sheetNodes.size(); ++i)
        {
            QDomElement sheet = sheetNodes.at(i).toElement();

            if (!sheet.isNull())
            {
                sheetIdToUserFriendlyNameMap[sheet.attribute(QStringLiteral(
                    "r:id"))] = sheet.attribute(QStringLiteral("name"));
            }
        }
    }
    else
    {
        setError(__FUNCTION__, "Can not open xl/workbook.xml file.");
        return {false, {}};
    }

    if (zip.setCurrentFile(QStringLiteral("xl/_rels/workbook.xml.rels")))
    {
        // Open file in archive.
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
            setError(__FUNCTION__, "File is corrupted.");
            return {false, {}};
        }
        zipFile.close();

        QDomElement root = xmlDocument.documentElement();
        QDomNodeList sheetNodes = root.childNodes();

        if (sheetNodes.size() <= 0)
        {
            setError(__FUNCTION__, "File is corrupted, no sheets in xml.");
            return {false, {}};
        }

        for (int i = 0; i < sheetNodes.size(); ++i)
        {
            QDomElement sheet = sheetNodes.at(i).toElement();

            if (!sheet.isNull())
            {
                QMap<QString, QString>::const_iterator iterator =
                    sheetIdToUserFriendlyNameMap.constFind(
                        sheet.attribute(QStringLiteral("Id")));

                if (sheetIdToUserFriendlyNameMap.constEnd() != iterator)
                {
                    sheets.append(
                        {*iterator,
                         "xl/" + sheet.attribute(QStringLiteral("Target"))});
                }
            }
        }
    }
    else
    {
        setError(__FUNCTION__,
                 "No file named xl/_rels/workbook.xml.rels in archive.");
        return {false, {}};
    }

    sheets_ = std::move(sheets);

    QStringList sheetsToReturn;
    for (const auto& [sheetName, sheetPath] : *sheets_)
        sheetsToReturn << sheetName;
    return {true, sheetsToReturn};
}

std::pair<bool, QList<std::pair<QString, QString>>> ImportXlsx::getSheets()
{
    if (!sheets_ && !getSheetNames().first)
        return {false, {}};
    return {true, *sheets_};
}

void ImportXlsx::setSheets(QList<std::pair<QString, QString>> sheets)
{
    sheets_ = std::move(sheets);
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

    QuaZip zip(&ioDevice_);

    if (!zip.open(QuaZip::mdUnzip))
    {
        setError(__FUNCTION__,
                 "Can not open zip file " + zip.getZipName() + ".");
        return {false, {}};
    }

    const unsigned int columnCount{columnCounts_.find(sheetName).value()};
    QStringList columnNames;
    // Loading column names is using Excel names names. Set 600 temporary.
    // const int columnsCount = EibleUtilities::getMaxExcelColumns();
    const QList<QByteArray> excelColNames =
        EibleUtilities::generateExcelColumnNames(columnCount);

    auto [sheetFound, sheetPath] = getSheetPath(sheetName);
    if (!sheetFound)
        return {false, {}};

    if (zip.setCurrentFile(sheetPath))
    {
        QuaZipFile zipFile(&zip);

        // Opening file.
        if (!zipFile.open(QIODevice::ReadOnly | QIODevice::Text))
        {
            setError(__FUNCTION__,
                     "Can not open file " + zipFile.getFileName() + ".");
            return {false, {}};
        }

        QXmlStreamReader xmlStreamReader;
        xmlStreamReader.setDevice(&zipFile);

        // Variable with actual type of data in cell (s, str, null).
        QString currentColType = QStringLiteral("s");

        // Go to first row.
        while (!xmlStreamReader.atEnd() &&
               xmlStreamReader.name() != "sheetData")
            xmlStreamReader.readNext();

        xmlStreamReader.readNext();
        xmlStreamReader.readNext();

        // Actual column number.
        int columnIndex = -1;
        QXmlStreamReader::TokenType lastToken = xmlStreamReader.tokenType();

        const QRegExp regExp(QLatin1String("[0-9]"));

        // Parse first row.
        while (!xmlStreamReader.atEnd() && xmlStreamReader.name() != "row")
        {
            // If we encounter start of cell content we add it to list.
            if (xmlStreamReader.name().toString() == QLatin1String("c") &&
                xmlStreamReader.tokenType() == QXmlStreamReader::StartElement)
            {
                columnIndex++;
                QString rowNumber(xmlStreamReader.attributes()
                                      .value(QLatin1String("r"))
                                      .toString());

                // If cells are missing add default name.
                while (excelColNames.indexOf(
                           rowNumber.remove(regExp).toUtf8()) > columnIndex)
                {
                    columnNames << emptyColName_;
                    columnIndex++;
                }
                // Remember column type.
                currentColType = xmlStreamReader.attributes()
                                     .value(QLatin1String("t"))
                                     .toString();
            }

            // If we encounter start of cell content than add it to list.
            if (!xmlStreamReader.atEnd() &&
                xmlStreamReader.name().toString() == QLatin1String("v") &&
                xmlStreamReader.tokenType() == QXmlStreamReader::StartElement)
            {
                if (currentColType == QLatin1String("s"))
                {
                    int value = xmlStreamReader.readElementText().toInt();
                    columnNames.push_back((*sharedStrings_)[value]);
                }
                else
                {
                    if (currentColType == QLatin1String("str"))
                        columnNames.push_back(
                            xmlStreamReader.readElementText());
                    else
                        columnNames.push_back(
                            xmlStreamReader.readElementText());
                }
            }

            // If we encounter empty cell than add it to list.
            if (xmlStreamReader.name().toString() == QLatin1String("c") &&
                xmlStreamReader.tokenType() == QXmlStreamReader::EndElement &&
                lastToken == QXmlStreamReader::StartElement)
                columnNames << emptyColName_;
            lastToken = xmlStreamReader.tokenType();
            xmlStreamReader.readNext();
        }
    }
    else
    {
        setError(__FUNCTION__,
                 "File named " + sheetName + " not found in archive.");
        return {false, {}};
    }

    while (columnNames.count() < static_cast<int>(columnCount))
        columnNames << emptyColName_;

    return {true, columnNames};
}

std::tuple<bool, std::optional<QList<int>>, std::optional<QList<int>>>
ImportXlsx::getStyles()
{
    QList<int> dateStyles;
    QList<int> allStyles;

    const QList predefinedExcelStylesForDates{14, 15, 16, 17, 22};
    // List of predefined excel styles for dates.
    dateStyles.append(predefinedExcelStylesForDates);

    QuaZip zip(&ioDevice_);
    if (!zip.open(QuaZip::mdUnzip))
    {
        setError(__FUNCTION__,
                 "Can not open zip file " + zip.getZipName() + ".");
        return {false, std::nullopt, std::nullopt};
    }

    // Load styles.
    if (zip.setCurrentFile(QStringLiteral("xl/styles.xml")))
    {
        // Open file in archive.
        QuaZipFile zipFile(&zip);
        if (!zipFile.open(QIODevice::ReadOnly | QIODevice::Text))
        {
            setError(__FUNCTION__, "Can not open file.");
            return {false, std::nullopt, std::nullopt};
        }

        // Create, set content and read DOM.
        QDomDocument xmlDocument(__FUNCTION__);
        if (!xmlDocument.setContent(zipFile.readAll()))
        {
            setError(__FUNCTION__, "Xml file is corrupted.");
            return {false, std::nullopt, std::nullopt};
        }
        zipFile.close();

        QDomElement root = xmlDocument.documentElement();
        QDomNodeList sheetNodes =
            root.firstChildElement(QStringLiteral("numFmts")).childNodes();

        for (int i = 0; i < sheetNodes.size(); ++i)
        {
            QDomElement sheet = sheetNodes.at(i).toElement();

            if (sheet.hasAttribute(QStringLiteral("numFmtId")) &&
                sheet.hasAttribute(QStringLiteral("formatCode")))  //&&
            // sheet.attribute("formatCode").contains("@"))
            {
                QString formatCode =
                    sheet.attribute(QStringLiteral("formatCode"));
                bool gotD = formatCode.contains(QStringLiteral("d"),
                                                Qt::CaseInsensitive);
                bool gotM = formatCode.contains(QStringLiteral("m"),
                                                Qt::CaseInsensitive);
                bool gotY = formatCode.contains(QStringLiteral("y"),
                                                Qt::CaseInsensitive);

                if ((gotD && gotY) || (gotD && gotM) || (gotM && gotY))
                {
                    dateStyles.push_back(
                        sheet.attribute(QStringLiteral("numFmtId")).toInt());
                }
            }
        }

        sheetNodes =
            root.firstChildElement(QStringLiteral("cellXfs")).childNodes();

        for (int i = 0; i < sheetNodes.size(); ++i)
        {
            QDomElement sheet = sheetNodes.at(i).toElement();

            if (!sheet.isNull() &&
                sheet.hasAttribute(QStringLiteral("numFmtId")))
            {
                allStyles.push_back(
                    sheet.attribute(QStringLiteral("numFmtId")).toInt());
            }
        }
    }
    else
    {
        setError(__FUNCTION__, "No file named xl/workbook.xml in archive.");
        return {false, std::nullopt, std::nullopt};
    }

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

    // Loading shared strings, it is separate file in archive with unique
    // table of all strings, in spreadsheet there are calls to this table.

    QuaZip zip(&ioDevice_);

    if (!zip.open(QuaZip::mdUnzip))
    {
        setError(__FUNCTION__,
                 "Can not open zip file " + zip.getZipName() + ".");
        return {false, {}};
    }

    QStringList sharedStrings;

    if (zip.setCurrentFile(QStringLiteral("xl/sharedStrings.xml")))
    {
        // Set variable.
        QuaZipFile zipFile(&zip);

        // Opening file.
        if (!zipFile.open(QIODevice::ReadOnly | QIODevice::Text))
        {
            setError(__FUNCTION__,
                     "Can not open file " + zipFile.getFileName() + ".");
            return {false, {}};
        }

        QXmlStreamReader xmlStreamReader;
        xmlStreamReader.setDevice(&zipFile);

        while (!xmlStreamReader.atEnd())
        {
            xmlStreamReader.readNext();

            // If 't' tag found add value to shared strings.
            if (xmlStreamReader.name() == "t")
                sharedStrings.append(xmlStreamReader.readElementText());
        }
        zipFile.close();
    }
    else
    {
        setError(__FUNCTION__, "No file xl/sharedStrings.xml in archive.");
        return {true, {}};
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

    QuaZip zip(&ioDevice_);

    if (!zip.open(QuaZip::mdUnzip))
    {
        setError(__FUNCTION__,
                 "Can not open zip file " + zip.getZipName() + ".");
        return {false, {}};
    }

    auto [sheetFound, sheetPath] = getSheetPath(sheetName);
    if (!sheetFound)
        return {false, {}};

    QuaZipFile zipFile;
    QXmlStreamReader xmlStreamReader;

    if (!openZipAndMoveToSecondRow(zip, sheetPath, zipFile, xmlStreamReader))
        return {false, {}};

    QVector<ColumnType> columnTypes;

    const QList<QByteArray> excelColNames =
        EibleUtilities::generateExcelColumnNames(
            EibleUtilities::getMaxExcelColumns());

    // Current column.
    int column = NOT_SET_COLUMN;
    int maxColumn = NOT_SET_COLUMN;

    // Current row.
    int rowCounter = 0;

    int charsToChopFromEndInCellName = 1;

    QXmlStreamAttributes xmlStreamAtrributes;

    // Optimization.
    const QString rowTag(QStringLiteral("row"));
    const QString cellTag(QStringLiteral("c"));
    const QString sheetDataTag(QStringLiteral("sheetData"));
    const QString sTag(QStringLiteral("s"));
    const QString strTag(QStringLiteral("str"));
    const QString rTag(QStringLiteral("r"));
    const QString tTag(QStringLiteral("t"));
    const QString vTag(QStringLiteral("v"));

    while (!xmlStreamReader.atEnd() &&
           0 != xmlStreamReader.name().compare(sheetDataTag))
    {
        // If start of row encountered than reset column counter add
        // increment row counter.
        if (0 == xmlStreamReader.name().compare(rowTag) &&
            xmlStreamReader.isStartElement())
        {
            column = NOT_SET_COLUMN;
            rowCounter++;
            double power = pow(DECIMAL_BASE, charsToChopFromEndInCellName);
            if (power <= rowCounter + 1)
                charsToChopFromEndInCellName++;
        }

        // When we encounter start of cell description.
        if (0 == xmlStreamReader.name().compare(cellTag) &&
            xmlStreamReader.isStartElement())
        {
            column++;

            QString stringToChop =
                xmlStreamReader.attributes().value(rTag).toString();
            int numberOfCharsToRemove =
                stringToChop.size() - charsToChopFromEndInCellName;
            int expectedIndexCurrentColumn = excelColNames.indexOf(
                stringToChop.left(numberOfCharsToRemove).toUtf8());

            // If cells are missing increment column number.
            while (expectedIndexCurrentColumn > column)
                column++;

            maxColumn = std::max(maxColumn, column);

            // If we encounter column outside expected grid we move to row end.
            if (expectedIndexCurrentColumn == NOT_SET_COLUMN)
            {
                xmlStreamReader.skipCurrentElement();
                continue;
            }

            if (column >= columnTypes.size())
                for (int i = columnTypes.size(); i <= column; ++i)
                    columnTypes.push_back(ColumnType::UNKNOWN);

            // If data format in column is unknown than read it.
            if (columnTypes.at(column) == ColumnType::UNKNOWN)
            {
                xmlStreamAtrributes = xmlStreamReader.attributes();
                QString value = xmlStreamAtrributes.value(tTag).toString();
                bool valueIsSTag = (0 == value.compare(sTag));
                if (valueIsSTag || 0 == value.compare(strTag))
                {
                    if (valueIsSTag)
                    {
                        xmlStreamReader.readNextStartElement();

                        if (0 == xmlStreamReader.name().compare(vTag) &&
                            xmlStreamReader.tokenType() ==
                                QXmlStreamReader::StartElement &&
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
                        xmlStreamAtrributes.value(sTag).toString();
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
                    QString value = xmlStreamAtrributes.value(tTag).toString();
                    bool valueIsSTag = (0 == value.compare(sTag));
                    if (valueIsSTag || 0 == value.compare(strTag))
                    {
                        if (valueIsSTag)
                        {
                            xmlStreamReader.readNextStartElement();

                            if (0 == xmlStreamReader.name().compare(vTag) &&
                                xmlStreamReader.tokenType() ==
                                    QXmlStreamReader::StartElement &&
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
                            xmlStreamAtrributes.value(sTag).toString();

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

bool ImportXlsx::openZipAndMoveToSecondRow(QuaZip& zip,
                                           const QString& sheetName,
                                           QuaZipFile& zipFile,
                                           QXmlStreamReader& xmlStreamReader)
{
    if (!zip.setCurrentFile(sheetName))
    {
        setError(__FUNCTION__,
                 "File named " + sheetName + " not found in archive.");
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

    bool secondRow = false;
    while (!xmlStreamReader.atEnd())
    {
        if (xmlStreamReader.name() == "row" &&
            xmlStreamReader.tokenType() == QXmlStreamReader::StartElement)
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

bool ImportXlsx::analyzeSheet(const QString& sheetName)
{
    QuaZip zip(&ioDevice_);

    if (!zip.open(QuaZip::mdUnzip))
    {
        setError(__FUNCTION__,
                 "Can not open zip file " + zip.getZipName() + ".");
        return false;
    }

    auto [sheetFound, sheetPath] = getSheetPath(sheetName);
    if (!sheetFound)
        return false;

    QuaZipFile zipFile;
    QXmlStreamReader xmlStreamReader;

    if (!openZipAndMoveToSecondRow(zip, sheetPath, zipFile, xmlStreamReader))
        return false;

    const QList<QByteArray> excelColNames =
        EibleUtilities::generateExcelColumnNames(
            EibleUtilities::getMaxExcelColumns());

    const QString rowTag(QStringLiteral("row"));
    const QString cellTag(QStringLiteral("c"));
    const QString rTag(QStringLiteral("r"));
    const QString sheetDataTag(QStringLiteral("sheetData"));

    unsigned int rowCounter = 0;
    int maxColumn = NOT_SET_COLUMN;
    int column = NOT_SET_COLUMN;
    int charsToChopFromEndInCellName = 1;

    while (!xmlStreamReader.atEnd() &&
           0 != xmlStreamReader.name().compare(sheetDataTag))
    {
        if (0 == xmlStreamReader.name().compare(rowTag) &&
            xmlStreamReader.isStartElement())
        {
            rowCounter++;
            column = NOT_SET_COLUMN;
            double power = pow(DECIMAL_BASE, charsToChopFromEndInCellName);
            if (power <= rowCounter + 1)
                charsToChopFromEndInCellName++;
        }

        if (0 == xmlStreamReader.name().compare(cellTag) &&
            xmlStreamReader.isStartElement())
        {
            column++;
            QString stringToChop =
                xmlStreamReader.attributes().value(rTag).toString();
            int numberOfCharsToRemove =
                stringToChop.size() - charsToChopFromEndInCellName;
            int expectedIndexCurrentColumn = excelColNames.indexOf(
                stringToChop.left(numberOfCharsToRemove).toUtf8());

            // If cells are missing increment column number.
            while (expectedIndexCurrentColumn > column)
                column++;

            maxColumn = std::max(maxColumn, column);
        }
        xmlStreamReader.readNextStartElement();
    }

    rowCounts_[sheetName] = rowCounter;
    columnCounts_[sheetName] = maxColumn + 1;

    return true;
}

std::pair<bool, QList<int>> ImportXlsx::getAllStyles()
{
    if (allStyles_)
        return {true, *allStyles_};
    bool success{false};
    std::tie(success, dateStyles_, allStyles_) = getStyles();
    return {success, *allStyles_};
}

void ImportXlsx::setAllStyles(QList<int> allStyles)
{
    allStyles_ = std::move(allStyles);
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

    QuaZip zip(&ioDevice_);

    if (!zip.open(QuaZip::mdUnzip))
    {
        setError(__FUNCTION__,
                 "Can not open zip file " + zip.getZipName() + ".");
        return {false, {}};
    }

    auto [sheetFound, sheetPath] = getSheetPath(sheetName);
    if (!sheetFound)
        return {false, {}};

    QList<QByteArray> excelColNames =
        EibleUtilities::generateExcelColumnNames(columnCount);

    QuaZipFile zipFile;
    QXmlStreamReader xmlStreamReader;

    openZipAndMoveToSecondRow(zip, sheetPath, zipFile, xmlStreamReader);

    // Actual column number.
    int column = NOT_SET_COLUMN;

    // Actual data type in cell (s, str, null).
    QString currentColType = QStringLiteral("s");
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
    int charsToChopFromEndInCellName = 1;

    // Optimization.
    const QString rowTag(QStringLiteral("row"));
    const QString cellTag(QStringLiteral("c"));
    const QString sheetDataTag(QStringLiteral("sheetData"));
    const QString sTag(QStringLiteral("s"));
    const QString vTag(QStringLiteral("v"));
    const QString rTag(QStringLiteral("r"));
    const QString tTag(QStringLiteral("t"));

    unsigned int lastEmittedPercent{0};
    while (!xmlStreamReader.atEnd() &&
           0 != xmlStreamReader.name().compare(sheetDataTag) &&
           rowCounter <= rowLimit)
    {
        // If start of row encountered than reset column counter add
        // increment row counter.
        if (0 == xmlStreamReader.name().compare(rowTag) &&
            xmlStreamReader.isStartElement())
        {
            column = NOT_SET_COLUMN;

            if (0 != rowCounter)
            {
                dataContainer[rowCounter - 1] = currentDataRow;
                currentDataRow = QVector<QVariant>(templateDataRow);

                double power = pow(DECIMAL_BASE, charsToChopFromEndInCellName);
                if (static_cast<unsigned int>(qRound(power)) <= rowCounter + 2)
                    charsToChopFromEndInCellName++;
            }

            rowCounter++;

            updateProgress(rowCounter, rowLimit, lastEmittedPercent);

            if (fillSamplesOnly && rowCounter > containerSize)
                break;
        }

        // When we encounter start of cell description.
        if (0 == xmlStreamReader.name().compare(cellTag) &&
            xmlStreamReader.isStartElement())
        {
            column++;

            QString stringToChop =
                xmlStreamReader.attributes().value(rTag).toString();
            int expectedIndexCurrentColumn = excelColNames.indexOf(
                stringToChop
                    .left(stringToChop.size() - charsToChopFromEndInCellName)
                    .toUtf8());

            // If cells missing than increment column number.
            while (expectedIndexCurrentColumn > column)
                column++;

            // If we encounter column outside expected grid we move to row end.
            if (expectedIndexCurrentColumn == NOT_SET_COLUMN)
            {
                xmlStreamReader.skipCurrentElement();
                continue;
            }

            // Remember cell type.
            currentColType =
                xmlStreamReader.attributes().value(tTag).toString();
            actualSTagValue =
                xmlStreamReader.attributes().value(sTag).toString();
        }

        // Insert data into table.
        if (!xmlStreamReader.atEnd() &&
            0 == xmlStreamReader.name().compare(vTag) &&
            xmlStreamReader.tokenType() == QXmlStreamReader::StartElement &&
            (fillSamplesOnly || !excludedColumns.contains(column)))
        {
            ColumnType format = columnTypes.at(column);

            switch (format)
            {
                case ColumnType::STRING:
                {
                    // Strings.
                    if (0 == currentColType.compare(sTag))
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
                    if (0 != currentColType.compare(sTag))
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

void ImportXlsx::setDateStyles(QList<int> dateStyles)
{
    dateStyles_ = std::move(dateStyles);
}

void ImportXlsx::setSharedStrings(QStringList sharedStrings)
{
    sharedStrings_ = std::move(sharedStrings);
}
