# About project
 Library for importing and exporting data. Created as a result of division of code of Volbx project and moving parts of it to independent library. Library contains:  
 + ExportData (base class for exporting classes)
 + ExportXlsx
 + ExportDsv
 + ImportSpreadsheet (base class for spreadsheet exporting classes)
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
## Import
[TODO]
### ImportSpreadsheet
[TODO]
### Xlsx
[TODO]
### Ods
[TODO]
## Export
[TODO]
### ExportData
[TODO]
### Dsv
[TODO]
### Xlsx
[TODO]