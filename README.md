## Table of content
- [About project](#about-project)
- [Building](#building)
- [Usage](#usage)
- [Classes](#classes)


# About project
 This is a library for importing and exporting data. Created as a result of the division of Volbx project code and moving parts of it to an independent library. The library contains classes:  
 + ExportData (base class for exporting)
 + ExportXlsx
 + ExportDsv
 + ImportSpreadsheet (base class for spreadsheet importing)
 + ImportOds
 + ImportXlsx  

The library is based on Qt 6 and C++17. It uses the Zlib and QuaZip projects.

# Building
Clone and use CMake directly or via Qt Creator. CMake **should**:
+ configure everything automatically,
+ compile library and create binaries.

**TIP 1**: remember to set properly `CMAKE_PREFIX_PATH` env variable. It should have a Qt installation path to let CMake use `find_package` command.  

**TIP 2**: make sure you install the `Core5Compat` module, which is part of Qt 6 as Quazip needs it.  

As a result of compilation, a dynamic lib should be created along with a headers' dir.

Check out my other project, Volbx to familiarize yourself with how to use it via CMake.

## Used tools and libs
| Tool |  Windows | Lubuntu |
| --- | --- | --- |
| OS version | 10 22H2 | 22.04 |
| GCC | 11.2.0 | 11.3.0 |
| CMake | 3.25.0 | 3.25.0 |
| Git | 2.38.1 | 2.34.1 |
| Qt | 6.5.2 | 6.5.2 |
| Qt Creator | 10.0.2 | 10.0.2 |
| Zlib | 1.3.1 | 1.3.1 |
| QuaZip | 1.4 | 1.4 |

# Usage
The easiest way is to check the examples' subproject, where you can find how to create and interact with each class included in this library.  
Alternatively, tests in a subproject can be checked. Usage can also be found in my other project called Volbx where classes from this library are used for exporting and importing data.

# Classes
## Classes used for importing data
### ImportSpreadsheet
Base class for spreadsheet-related import classes. The following pure virtual methods need to be implemented when creating a new derived class:
+ `getSheetNames()`
+ `getColumnNames()`
+ `getColumnTypes()`
+ `getCount()`
+ `getLimitedData()`
+ `retrieveColumnNames()`
+ `retrieveRowCountAndColumnTypes()`

Emits signal `progressPercentChanged` during loading data.

Importing can be done from objects with the interface QIODevice: files on disk, files from resources, QBuffer and more.
### ImportXlsx
Class used to import data from .xlsx files. Basic usage:
```cpp
QFile xlsxFile("example.xlsx");
ImportXlsx importXlsx(xlsxFile);
auto [success, xlsxData] = importXlsx.getData("SheetName", {});
```
Check `example` and `test` subprojects for more advanced use cases.  
Take note that .xlsx files store strings separately. ImportXlsx follows that convention and provides `getSharedStrings()` to retrieve those.
### ImportOds
Class used to import data from .ods files. Basic usage:
```cpp
QFile odsFile("example.ods");
ImportOds importOds(odsFile);
auto [success, odsData] = importOds.getData("SheetName", {});
```
Check `example` and `test` subprojects for more advanced use cases.  
## Classes used for exporting data
### ExportData
Base class for export-related classes. The following pure virtual methods need to be implemented when creating a new derived class:
+ `writeContent()`
+ `getEmptyContent()`
+ `generateHeaderContent()`
+ `generateRowContent()`

Emits signal `progressPercentChanged` during loading data.

Data can be exported to objects with the interface QIODevice. It can be not only a QFile but also QBuffer. Example of exporting data into memory:
```cpp
QByteArray exportedZipBuffer;
QBuffer exportedZip(&exportedZipBuffer);
exportedZip.open(QIODevice::WriteOnly);
QTableWidget tableWidget;
ExportXlsx exportXlsx;
exportXlsx.exportView(tableWidget, exportedZip);
```
### ExportDsv
Class for exporting data to DSV (Delimiter Separated Values) files. The delimiter is set in the constructor and can be any character (comma, tab, semicolon, ...). CSV or TSV files can be created using this class.   
Three additional methods can be used to customize the output:
+ `setDateFormat(Qt::DateFormat)` - for setting the output date format to a given one,
+ `setDateFormat(QString)` - for setting the output date format to one given as a string (for example, "dd/MM/yy"),
+ `setNumbersLocale(QLocale)`- for setting the locale used for output numbers in case a different one from the default one is needed.  

Basic usage:
```cpp
QFile outFile("example.csv");
outFile.open(QIODevice::WriteOnly);
QTableWidget tableWidget;
ExportDsv exportDsv(',');
exportDsv.exportView(tableWidget, outFile);
```
### ExportXlsx
Class for exporting data to .xlsx files.  
Basic usage:
```cpp
QFile outFile("example.xlsx");
outFile.open(QIODevice::WriteOnly);
QTableWidget tableWidget;
ExportXlsx exportXlsx;
exportXlsx.exportView(tableWidget, outFile);
```
XLSX files are created using the template file `template.xlsx` included in the project via the resource collection file `resources.qrc`.