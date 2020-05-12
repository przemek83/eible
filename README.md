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
### ImportXlsx
Class used to import data from .xlsx files. Basic usage:
```cpp
QFile xlsxTestFile("example.xlsx);
ImportXlsx importXlsx(xlsxTestFile);
auto [success, xlsxData] = importXlsx.getDate("SheetName", {});
```
Check `example` and `test` subprojects for more advanced usage cases.  
Take note that .xlsx files stores strings separately. ImportXlsx follows that convention and provides `getSharedStrings()` to retrieve those.
### Ods
[TODO]
## Classes used for exporting data
[TODO]
### ExportData
[TODO]
### Dsv
[TODO]
### Xlsx
[TODO]