# About project
 Library for importing and exporting data. Created as a result of division of code of Volbx project and moving parts of it to independent library. Library contains classes:  
 + ExportData (base class for exporting)
 + ExportXlsx
 + ExportDsv
 + ImportSpreadsheet (base class for spreadsheet importing)
 + ImportOds
 + ImportXlsx
  
# Building
Clone and use Cmake directly or via QtCreator. Cmake **should**:
+ configure everything automatically,
+ compile library and create binaries.

**TIP**: remember to set properly `CMAKE_PREFIX_PATH` env variable. It should have Qt installation path to let Cmake `find_package` command work.  

As a result of compilation dynamic lib should be created along with headers dir.

To use it as external project via Cmake you may check how it is done in my other project called Volbx.

# Usage
Easiest way is to check examples subproject where you can find how to create and interact with each class included in this library.  
Alternatively tests subproject can be checked. Usage also can be found in my other project called Volbx where classes from this library are used.

# Classes
## Classes used for importing data
### ImportSpreadsheet
Base class for spreadsheet related import classes. Following pure virtual methods need to be implemented when creating new derived class:
+ `getSheetNames()`
+ `getColumnNames()`
+ `getColumnTypes()`
+ `getColumnCount()`
+ `getRowCount()`
+ `getLimitedData()`
+ `retrieveColumnNames()`
+ `retrieveRowCountAndColumnTypes()`

Emits signal `progressPercentChanged()` during loading data.
### ImportXlsx
Class used to import data from .xlsx files. Basic usage:
```cpp
QFile xlsxFile("example.xlsx");
ImportXlsx importXlsx(xlsxFile);
auto [success, xlsxData] = importXlsx.getData("SheetName", {});
```
Check `example` and `test` subprojects for more advanced usage cases.  
Take note that .xlsx files stores strings separately. ImportXlsx follows that convention and provides `getSharedStrings()` to retrieve those.
### ImportOds
Class used to import data from .ods files. Basic usage:
```cpp
QFile odsFile("example.ods");
ImportOds importOds(odsFile);
auto [success, odsData] = importOds.getData("SheetName", {});
```
Check `example` and `test` subprojects for more advanced usage cases.  
## Classes used for exporting data
### ExportData
Base class for export related classes. Following pure virtual methods need to be implemented when creating new derived class:
+ `writeContent()`
+ `getEmptyContent()`
+ `generateHeaderContent()`
+ `generateRowContent()`

Emits signal `progressPercentChanged` during loading data.
### Dsv
Class for exporting data to DSV (Delimiter Separated Values) files. Delimiter is set in constructor and can be any char (comma, tab, semicolon, ...). Via that class files of type CSV or TSV can be created. Three additional method can be used to customize output:
+ `setDateFormat(Qt::DateFormat)` - for setting output date format to Qt defined one,
+ `setDateFormat(QString)` - for setting output date format to one given as string (for example "dd/MM/yy"),
+ `setNumbersLocale(QLocale)`- for setting locale used for output numbers in case different than default one for application need to be used.

Basic usage:
```cpp
ExportDsv exportDsv(',');
QFile outFile("example.csv");
outFile.open(QIODevice::WriteOnly);
QTableWidget tableWidget;
bool success = exportDsv.exportView(tableWidget, outFile);
```
### Xlsx
[TODO]