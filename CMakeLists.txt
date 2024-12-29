cmake_minimum_required(VERSION 3.16)

project(eible
        LANGUAGES CXX
        VERSION 1.1.0
        DESCRIPTION "Export and import library"
        HOMEPAGE_URL "https://github.com/przemek83/eible")

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

option(CODE_COVERAGE "Collect coverage" OFF)

if(CODE_COVERAGE)
    set(CMAKE_C_FLAGS "${CMAKE_C_FLAGS} -O0 --coverage")
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -O0 --coverage")
endif()

find_package(Qt6 REQUIRED COMPONENTS Widgets Core Test Xml Core5Compat)
qt_standard_project_setup()

include(GNUInstallDirs)
include(CMakePackageConfigHelpers)

include(ZlibConfig.cmake)
include(QuaZipConfig.cmake)

set(${PROJECT_NAME}_SOURCES 
    src/Utilities.h
    src/Utilities.cpp
    src/ExportXlsx.cpp
    src/ExportDsv.cpp
    src/ExportData.cpp
    src/ImportSpreadsheet.cpp
    src/ImportXlsx.cpp
    src/ImportOds.cpp

)
qt_add_resources(${PROJECT_NAME}_RESOURCES res/resources.qrc)

set(PUBLIC_HEADERS
    include/${PROJECT_NAME}/eible_global.h
    include/${PROJECT_NAME}/ExportXlsx.h
    include/${PROJECT_NAME}/ExportDsv.h
    include/${PROJECT_NAME}/ExportData.h
    include/${PROJECT_NAME}/ImportSpreadsheet.h
    include/${PROJECT_NAME}/ImportXlsx.h
    include/${PROJECT_NAME}/ImportOds.h
    include/${PROJECT_NAME}/ColumnType.h
)

add_library(${PROJECT_NAME} SHARED ${${PROJECT_NAME}_SOURCES} ${PUBLIC_HEADERS} ${${PROJECT_NAME}_RESOURCES})

target_include_directories(${PROJECT_NAME}
  PUBLIC
    $<BUILD_INTERFACE:${PROJECT_SOURCE_DIR}/include>
    $<INSTALL_INTERFACE:${CMAKE_INSTALL_INCLUDEDIR}>
  PRIVATE
    $<BUILD_INTERFACE:${PROJECT_SOURCE_DIR}/src>)

set_target_properties(${PROJECT_NAME} PROPERTIES
    VERSION ${PROJECT_VERSION}
    SOVERSION ${PROJECT_VERSION_MAJOR}
    PUBLIC_HEADER "${PUBLIC_HEADERS}")

configure_file(${PROJECT_NAME}.pc.in ${PROJECT_NAME}.pc @ONLY)

target_compile_definitions(${PROJECT_NAME} PRIVATE EIBLE_LIBRARY)

install(TARGETS ${PROJECT_NAME} 
    LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
    ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR}
    RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR}
    PUBLIC_HEADER DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}/${PROJECT_NAME})

install(FILES ${CMAKE_BINARY_DIR}/${PROJECT_NAME}.pc
    DESTINATION ${CMAKE_INSTALL_DATAROOTDIR}/pkgconfig)

target_link_libraries(${PROJECT_NAME} PRIVATE Qt6::Core Qt6::Widgets Qt6::Xml Qt6::Core5Compat QuaZip)

if(CMAKE_PROJECT_NAME STREQUAL PROJECT_NAME)
    add_subdirectory(examples)
    enable_testing()
    add_subdirectory(tests)
endif()
