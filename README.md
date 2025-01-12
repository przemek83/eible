[![Build & test](https://github.com/przemek83/eible/actions/workflows/buld-and-test.yml/badge.svg)](https://github.com/przemek83/eible/actions/workflows/buld-and-test.yml)
[![CodeQL](https://github.com/przemek83/eible/actions/workflows/codeql.yml/badge.svg)](https://github.com/przemek83/eible/actions/workflows/codeql.yml)
[![codecov](https://codecov.io/gh/przemek83/eible/graph/badge.svg?token=DLFS4H0283)](https://codecov.io/gh/przemek83/eible)

[![Quality Gate Status](https://sonarcloud.io/api/project_badges/measure?project=przemek83_eible&metric=alert_status)](https://sonarcloud.io/summary/new_code?id=przemek83_eible)
[![Bugs](https://sonarcloud.io/api/project_badges/measure?project=przemek83_eible&metric=bugs)](https://sonarcloud.io/summary/new_code?id=przemek83_eible)
[![Code Smells](https://sonarcloud.io/api/project_badges/measure?project=przemek83_eible&metric=code_smells)](https://sonarcloud.io/summary/new_code?id=przemek83_eible)
[![Coverage](https://sonarcloud.io/api/project_badges/measure?project=przemek83_eible&metric=coverage)](https://sonarcloud.io/summary/new_code?id=przemek83_eible)
[![Duplicated Lines (%)](https://sonarcloud.io/api/project_badges/measure?project=przemek83_eible&metric=duplicated_lines_density)](https://sonarcloud.io/summary/new_code?id=przemek83_eible)

## Table of contents
- [About project](#about-project)
- [Getting Started](#getting-started)
   * [Prerequisites](#prerequisites)
   * [Building](#building)
   * [CMake integration](#cmake-integration)
- [Built with](#built-with)
- [Usage](#usage)
- [Classes](#classes)
   * [Classes used for importing data](#classes-used-for-importing-data)
      + [ImportSpreadsheet](#importspreadsheet)
      + [ImportXlsx](#importxlsx)
      + [ImportOds](#importods)
   * [Classes used for exporting data](#classes-used-for-exporting-data)
      + [ExportData](#exportdata)
      + [ExportDsv](#exportdsv)
      + [ExportXlsx](#exportxlsx)
- [Testing](#testing)
- [Licensing](#licensing)

## About project
 This is a library for importing and exporting data. Created as a result of the division of Volbx project code and moving parts of it to an independent library. The library contains classes:  
 + ExportData (base class for exporting)
 + ExportXlsx
 + ExportDsv
 + ImportSpreadsheet (base class for spreadsheet importing)
 + ImportOds
 + ImportXlsx  

The library is based on Qt 6 and uses the Zlib and QuaZip projects for managing archives.

## Getting Started
This section describes briefly how to setup the environment and build the project.

### Prerequisites
Qt in version 6.5 or greater, a C++ compiler with C++17 support as a minimum, and CMake 3.16+. 

### Building
Clone and use CMake directly or via any IDE supporting it. CMake should:
- configure everything automatically,
- compile and create binaries.

As a result of compilation, binary for simulations and binary for testing should be created.

**TIP**: Remember to set properly the `CMAKE_PREFIX_PATH` env variable. It should have a Qt installation path to let CMake `find_package` command work.  

**TIP**: Make sure you install the `Core5Compat` module, which is part of Qt 6 as QuaZip needs it.  

### CMake integration
Use `FetchContent` CMake module in your project:
```cmake
include(FetchContent)

FetchContent_Declare(
   eible
   GIT_REPOSITORY https://github.com/przemek83/eible
   GIT_TAG v1.2.0
)

FetchContent_MakeAvailable(eible)
```
From that moment, eible library can be used in the `target_link_libraries` command:
```cmake
add_executable(${PROJECT_NAME} ${SOURCES})
target_link_libraries(${PROJECT_NAME} shared eible)
```
Check my other project `Volbx` for real world CMake integration.

## Built with
| |  Windows | Windows | Ubuntu |
| --- | --- | --- | --- | 
| OS version | 10 22H2 | 10 22H2 | 24.04 |
| compiler | GCC 13.1.0 | MSVC 19.29 | GCC 13.2.0 |
| CMake | 3.30.2 | 3.30.2 |3.28.3 |
| Git | 2.46.0 | 2.46.0 | 2.43.0 |
| Qt | 6.5.2 | 6.5.2 | 6.5.2 |
| Qwt | 6.3.1 | 6.3.1 | 6.3.1 |

## Usage
The easiest way is to check the examples' subproject, where you can find how to create and interact with each class included in this library.  
Alternatively, tests in a subproject can be checked. Usage can also be found in my other project called Volbx where classes from this library are used for exporting and importing data.

## Classes
### Classes used for importing data
#### ImportSpreadsheet
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
#### ImportXlsx
Class used to import data from .xlsx files. Basic usage:
```cpp
QFile xlsxFile("example.xlsx");
ImportXlsx importXlsx(xlsxFile);
auto [success, xlsxData] = importXlsx.getData("SheetName", {});
```
Check `example` and `test` subprojects for more advanced use cases.  
Take note that .xlsx files store strings separately. ImportXlsx follows that convention and provides `getSharedStrings()` to retrieve those.
#### ImportOds
Class used to import data from .ods files. Basic usage:
```cpp
QFile odsFile("example.ods");
ImportOds importOds(odsFile);
auto [success, odsData] = importOds.getData("SheetName", {});
```
Check `example` and `test` subprojects for more advanced use cases.  
### Classes used for exporting data
#### ExportData
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
#### ExportDsv
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
#### ExportXlsx
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

## Testing
For testing purposes, the Qt Test framework is used. Build the project first. Make sure that the `eible-tests` target is built. Modern IDEs supporting CMake also support running tests with monitoring of failures. But in case you would like to run it manually, go to the `build/tests` directory, where the⁣ binary `eible-tests` should be available. Launching it should produce the following output on Linux:
Example run:
```
$ ./eible-tests
********* Start testing of ExportXlsxTest *********
Config: Using QtTest library 6.5.2, Qt 6.5.2 (x86_64-little_endian-lp64 shared (dynamic) release build; by GCC 10.3.1 20210422 (Red Hat 10.3.1-1)), ubuntu 24.04
PASS   : ExportXlsxTest::initTestCase()
PASS   : ExportXlsxTest::testExportingEmptyTable()
PASS   : ExportXlsxTest::testExportingHeadersOnly()

(...)

PASS   : ImportOdsTest::testInvalidSheetName()
PASS   : ImportOdsTest::testDamagedFile()
PASS   : ImportOdsTest::cleanupTestCase()
Totals: 71 passed, 0 failed, 1 skipped, 0 blacklisted, 142ms
********* Finished testing of ImportOdsTest *********

```
As an alternative, CTest can be used to run tests from the `build/tests` directory:
```
$ ctest
Test project <path>/eible/build/tests
    Start 1: eible-tests
1/1 Test #1: eible-tests ......................   Passed    0.29 sec

100% tests passed, 0 tests failed out of 1

Total Test time (real) =   0.29 sec

```

Performance tests for importing and exporting are disabled by default. Search for `QSKIP` and comment lines with that macro to activate it.

## Licensing
Eible library is published under a LGPL license.

The project uses the following software:
| Name | License | Home | Description |
| --- | --- | --- | --- |
| Qt | LGPLv3 | https://www.qt.io/| cross-platform application development framework |
| Zlib | Zlib license | https://www.zlib.net/ | compression library |
| QuaZip | LGPLv2.1 with static linking exception | https://github.com/stachenov/quazip | C++ wrapper for Minizip using Qt library |
