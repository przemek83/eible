## Table of content
- [About project](#about-project)
- [Building](#building)
- [Usage](#usage)
- [Classes](#classes)


# About project
 Library for importing and exporting data. Created as a result of division of  Volbx project code and moving parts of it to independent library. Library contains classes:  
 + ExportData (base class for exporting)
 + ExportXlsx
 + ExportDsv
 + ImportSpreadsheet (base class for spreadsheet importing)
 + ImportOds
 + ImportXlsx  

Library is based on Qt 5 and C++17. It uses Zlib and QuaZip projects.

# Building
Clone and use Cmake directly or via QtCreator. Cmake **should**:
+ configure everything automatically,
+ compile library and create binaries.

**TIP**: remember to set properly `CMAKE_PREFIX_PATH` env variable. It should have Qt installation path to let Cmake command `find_package` command.  

As a result of compilation dynamic lib should be created along with headers dir.

Check my other project called Volbx to familiarize yourself how to use it via CMake.

## Used tools and libs
| Tool |  Windows | Lubuntu |
| --- | --- | --- |
| OS version | 10 22H2 | 22.04 |
| GCC | 8.1.0 | 9.5.0 |
| CMake | 3.25.0 | 3.25.0 |
| Git | 2.38.1 | 2.34.1 |
| Svn | 1.14.2 | 1.14.1 |
| Qt | 5.15.2 | 5.15.2 |
| QtCreator | 9.0.0 |9.0.0 |
| Qwt | 6.1.6 | 6.1.6 |

# Usage
Easiest way is to check examples subproject where you can find how to create and interact with each class included in this library.  
Alternatively tests subproject can be checked. Usage also can be found in my other project called Volbx where classes from this library are used for exporting and importing data.

# Classes
## Classes used for importing data
### ImportSpreadsheet
Base class for spreadsheet related import classes. Following pure virtual methods need to be implemented when creating new derived class:
+ `getSheetNames()`
+ `getColumnNames()`
+ `getColumnTypes()`
+ `getCount()`
+ `getLimitedData()`
+ `retrieveColumnNames()`
+ `retrieveRowCountAndColumnTypes()`

Emits signal `progressPercentChanged` during loading data.

Importing can be done from object with interface QIODevice: file on disk, file from resources, QBuffer and more.
### ImportXlsx
Class used to import data from .xlsx files. Basic usage:
```cpp
QFile xlsxFile("example.xlsx");
ImportXlsx importXlsx(xlsxFile);
auto [success, xlsxData] = importXlsx.getData("SheetName", {});
```
Check `example` and `test` subprojects for more advanced use cases.  
Take note that .xlsx files stores strings separately. ImportXlsx follows that convention and provides `getSharedStrings()` to retrieve those.
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
Base class for export related classes. Following pure virtual methods need to be implemented when creating new derived class:
+ `writeContent()`
+ `getEmptyContent()`
+ `generateHeaderContent()`
+ `generateRowContent()`

Emits signal `progressPercentChanged` during loading data.

Data can be exported to object with interface QIODevice. It can be not only QFile but also QBuffer. Example of exporting data into memory:
```cpp
QByteArray exportedZipBuffer;
QBuffer exportedZip(&exportedZipBuffer);
exportedZip.open(QIODevice::WriteOnly);
QTableWidget tableWidget;
ExportXlsx exportXlsx;
exportXlsx.exportView(tableWidget, exportedZip);
```
### ExportDsv
Class for exporting data to DSV (Delimiter Separated Values) files. Delimiter is set in constructor and can be any char (comma, tab, semicolon, ...). CSV or TSV files can be created using this class.   
Three additional method can be used to customize output:
+ `setDateFormat(Qt::DateFormat)` - for setting output date format to Qt defined one,
+ `setDateFormat(QString)` - for setting output date format to one given as string (for example "dd/MM/yy"),
+ `setNumbersLocale(QLocale)`- for setting locale used for output numbers in case different than default one is needed.  

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
Xlsx files are created using template file `template.xlsx` included into project via resource collection file `resources.qrc`.
