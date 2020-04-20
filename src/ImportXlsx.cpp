#include "ImportXlsx.h"

#include <quazip5/quazip.h>
#include <quazip5/quazipfile.h>
#include <QIODevice>
#include <QMap>
#include <QSet>
#include <QXmlStreamReader>
#include <QtXml/QDomDocument>

#include "EibleUtilities.h"

ImportXlsx::ImportXlsx(QIODevice& ioDevice) : ImportSpreadsheet(ioDevice) {}

std::pair<bool, QMap<QString, QString> > ImportXlsx::getSheetList()
{
    QuaZip zip(&ioDevice_);

    if (!zip.open(QuaZip::mdUnzip))
    {
        setError(__FUNCTION__,
                 "Can not open zip file " + zip.getZipName() + ".");
        return {false, {}};
    }

    QMap<QString, QString> sheetIdToUserFriendlyNameMap;
    QMap<QString, QString> sheetToFileMapInZip;

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
                    sheetToFileMapInZip[*iterator] =
                        "xl/" + sheet.attribute(QStringLiteral("Target"));
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

    return {true, sheetToFileMapInZip};
}

std::pair<bool, QStringList> ImportXlsx::getColumnList(
    const QString& sheetPath, const QHash<QString, int>& sharedStrings)
{
    QuaZip zip(&ioDevice_);

    if (!zip.open(QuaZip::mdUnzip))
    {
        setError(__FUNCTION__,
                 "Can not open zip file " + zip.getZipName() + ".");
        return {false, {}};
    }

    QStringList columnList;
    // Loading column names is using Excel names names. Set 600 temporary.
    const int columnsCount = EibleUtilities::getMaxExcelColumns();
    QStringList excelColNames =
        EibleUtilities::generateExcelColumnNames(columnsCount);

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
        {
            xmlStreamReader.readNext();
        }

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
                while (excelColNames.indexOf(rowNumber.remove(regExp)) >
                       columnIndex)
                {
                    columnList << emptyColName_;
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
                    columnList.push_back(sharedStrings.key(value));
                }
                else
                {
                    if (currentColType == QLatin1String("str"))
                    {
                        columnList.push_back(xmlStreamReader.readElementText());
                    }
                    else
                    {
                        columnList.push_back(xmlStreamReader.readElementText());
                    }
                }
            }

            // If we encounter empty cell than add it to list.
            if (xmlStreamReader.name().toString() == QLatin1String("c") &&
                xmlStreamReader.tokenType() == QXmlStreamReader::EndElement &&
                lastToken == QXmlStreamReader::StartElement)
            {
                columnList << emptyColName_;
            }
            lastToken = xmlStreamReader.tokenType();
            xmlStreamReader.readNext();
        }
    }
    else
    {
        setError(__FUNCTION__,
                 "File named " + sheetPath + " not found in archive.");
        return {false, {}};
    }

    return {true, columnList};
}

std::tuple<bool, QList<int>, QList<int> > ImportXlsx::getStyles()
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
        return {false, {}, {}};
    }

    // Load styles.
    if (zip.setCurrentFile(QStringLiteral("xl/styles.xml")))
    {
        // Open file in archive.
        QuaZipFile zipFile(&zip);
        if (!zipFile.open(QIODevice::ReadOnly | QIODevice::Text))
        {
            setError(__FUNCTION__, "Can not open file.");
            return {false, {}, {}};
        }

        // Create, set content and read DOM.
        QDomDocument xmlDocument(__FUNCTION__);
        if (!xmlDocument.setContent(zipFile.readAll()))
        {
            setError(__FUNCTION__, "Xml file is corrupted.");
            return {false, {}, {}};
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
        return {false, {}, {}};
    }

    return {true, dateStyles, allStyles};
}

std::pair<bool, QStringList> ImportXlsx::getSharedStrings()
{
    // Loading shared strings, it is separate file in archive with unique table
    // of all strings, in spreadsheet there are calls to this table.

    QStringList sharedStrings;
    QuaZip zip(&ioDevice_);

    if (!zip.open(QuaZip::mdUnzip))
    {
        setError(__FUNCTION__,
                 "Can not open zip file " + zip.getZipName() + ".");
        return {false, {}};
    }

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

    return {true, sharedStrings};
}

std::pair<bool, QVector<DataFormat> > ImportXlsx::getColumnDataFormats()
{
    return {false, {}};
}
