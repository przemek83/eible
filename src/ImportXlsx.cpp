#include "ImportXlsx.h"

#include <quazip5/quazip.h>
#include <quazip5/quazipfile.h>
#include <QIODevice>
#include <QMap>

#include <QtXml/QDomDocument>

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
