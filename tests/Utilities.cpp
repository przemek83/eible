#include "Utilities.h"

namespace Utilities
{
QByteArray composeXlsxSheet(const QString& sheetData)
{
    const std::string sheetTopPart{
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
)"
        R"(<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" )"
        R"(xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" )"
        R"(xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" )"
        R"(mc:Ignorable="x14ac" )"
        R"(xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">)"
        R"(<dimension ref="A1"/>)"
        R"(<sheetViews>)"
        R"(<sheetView tabSelected="1" workbookViewId="0"/>)"
        R"(</sheetViews>)"
        R"(<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>)"
        R"(<cols>)"
        R"(<col min="1" max="32" width="14.42578125" customWidth="1"/>)"
        R"(<col min="23" max="23" width="10.5703125"  customWidth="1"/>)"
        R"(<col min="24" max="24" width="11.5703125"  customWidth="1"/>)"
        R"(</cols>)"};

    const std::string sheetBottomPart{
        R"(<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" )"
        R"(footer="0.3"/></worksheet>)"};

    QString sheet;
    sheet.append(QString::fromStdString(sheetTopPart));
    sheet.append(sheetData);
    sheet.append(QString::fromStdString(sheetBottomPart));
    return sheet.toUtf8();
}
};  // namespace Utilities
