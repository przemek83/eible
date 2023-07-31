cmake_minimum_required(VERSION 3.16)

project(eible
        LANGUAGES CXX
        VERSION 1.1.0
        DESCRIPTION "Export and import library"
        HOMEPAGE_URL "https://github.com/przemek83/eible")

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

find_package(Qt6 REQUIRED COMPONENTS Widgets Core Test Xml Core5Compat)
qt_standard_project_setup()

set(THIRD_PARTY_LIBS zlib quazip)
foreach(library ${THIRD_PARTY_LIBS})
    configure_file(CMakeLists.txt.${library} ${library}/CMakeLists.txt)
    execute_process(COMMAND ${CMAKE_COMMAND} -G ${CMAKE_GENERATOR} .
            WORKING_DIRECTORY ${CMAKE_BINARY_DIR}/${library})
    execute_process(COMMAND ${CMAKE_COMMAND} --build .
            WORKING_DIRECTORY ${CMAKE_BINARY_DIR}/${library} )
endforeach(library)

set(ZLIB_ROOT ${CMAKE_BINARY_DIR}/zlib/install)
find_package(ZLIB REQUIRED)
list(APPEND CMAKE_PREFIX_PATH "${CMAKE_BINARY_DIR}/quazip/install/")
find_package(QuaZip-Qt6 CONFIG REQUIRED)

set(SOURCES
    src/EibleUtilities.cpp
    src/ExportXlsx.cpp
    src/ExportDsv.cpp
    src/ExportData.cpp
    src/ImportSpreadsheet.cpp
    src/ImportXlsx.cpp
    src/ImportOds.cpp
)
qt_add_resources(SOURCES res/resources.qrc)

set(HEADERS
    src/eible_global.h
    src/EibleUtilities.h
    src/ExportXlsx.h
    src/ExportDsv.h
    src/ExportData.h
    src/ImportSpreadsheet.h
    src/ImportXlsx.h
    src/ImportOds.h
    src/ColumnType.h
 )

add_library(${PROJECT_NAME} SHARED ${SOURCES} ${HEADERS})

set(PUBLIC_HEADERS
    src/eible_global.h
    src/ExportXlsx.h
    src/ExportDsv.h
    src/ExportData.h
    src/ImportSpreadsheet.h
    src/ImportXlsx.h
    src/ImportOds.h
    src/ColumnType.h
)

set_target_properties(${PROJECT_NAME} PROPERTIES
    VERSION ${PROJECT_VERSION}
    SOVERSION 1
    PUBLIC_HEADER "${PUBLIC_HEADERS}"
)

target_include_directories(${PROJECT_NAME} PUBLIC src)
target_link_libraries(${PROJECT_NAME} PRIVATE Qt6::Core Qt6::Widgets Qt6::Xml Qt6::Core5Compat QuaZip::QuaZip)

target_compile_definitions(${PROJECT_NAME} PRIVATE EIBLE_LIBRARY)

include(GNUInstallDirs)
install(TARGETS ${PROJECT_NAME}
    LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
    PUBLIC_HEADER DESTINATION ${CMAKE_INSTALL_INCLUDEDIR})

add_subdirectory(examples)

enable_testing()
add_subdirectory(tests)
