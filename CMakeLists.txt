cmake_minimum_required(VERSION 3.8)

project(eible LANGUAGES CXX VERSION 0.0.1 DESCRIPTION "Export import library")

set(CMAKE_INCLUDE_CURRENT_DIR ON)
set(CMAKE_AUTOUIC ON)
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

find_package(Qt5 COMPONENTS Widgets Core Test REQUIRED)
add_definitions( ${QT_DEFINITIONS} )

set(THIRD_PARTY_LIBS zlib quazip)
foreach(file_i ${THIRD_PARTY_LIBS})
    configure_file(CMakeLists.txt.${file_i} ${file_i}/CMakeLists.txt)
    execute_process(COMMAND ${CMAKE_COMMAND} -G ${CMAKE_GENERATOR} .
            WORKING_DIRECTORY ${CMAKE_BINARY_DIR}/${file_i})
    execute_process(COMMAND ${CMAKE_COMMAND} --build .
            WORKING_DIRECTORY ${CMAKE_BINARY_DIR}/${file_i} )
endforeach(file_i)

include_directories(${CMAKE_BINARY_DIR}/zlib/install/include)
include_directories(${CMAKE_BINARY_DIR}/quazip/install/include)

link_directories(${CMAKE_BINARY_DIR}/zlib/install/bin)
link_directories(${CMAKE_BINARY_DIR}/quazip/install/lib)

# copy dlls for easy launch
file(GLOB_RECURSE DLLS "${CMAKE_BINARY_DIR}/*.dll" )
file(COPY ${DLLS} DESTINATION ${CMAKE_BINARY_DIR}/examples)
file(COPY ${DLLS} DESTINATION ${CMAKE_BINARY_DIR}/tests)

qt5_add_resources(RESOURCES resources.qrc)
add_custom_target(generated_qrc DEPENDS ${RESOURCES})

add_library(${PROJECT_NAME} SHARED
  src/eible_global.h
  src/EibleUtilities.cpp
  src/EibleUtilities.h
  src/ExportXlsx.cpp
  src/ExportXlsx.h
  qrc_resources.cpp
)

add_dependencies(${PROJECT_NAME} generated_qrc)

set_target_properties(${PROJECT_NAME} PROPERTIES VERSION ${PROJECT_VERSION})
set_target_properties(${PROJECT_NAME} PROPERTIES SOVERSION 1)

set_property(TARGET ${PROJECT_NAME} PROPERTY PUBLIC_HEADER
    src/eible_global.h
    src/ExportXlsx.h
    src/EibleUtilities.h)
target_include_directories(${PROJECT_NAME} PUBLIC src)
include(GNUInstallDirs)

target_link_libraries(${PROJECT_NAME} PRIVATE Qt5::Core Qt5::Widgets)
target_link_libraries(${PROJECT_NAME} PRIVATE optimized quazip5)
target_link_libraries(${PROJECT_NAME} PRIVATE debug quazip5d)

target_compile_definitions(${PROJECT_NAME} PRIVATE EIBLE_LIBRARY)

add_custom_command(TARGET ${PROJECT_NAME} POST_BUILD
                   COMMAND ${CMAKE_COMMAND} -E copy
                           ${CMAKE_BINARY_DIR}/libeible.dll
                           ${CMAKE_BINARY_DIR}/examples/libeible.dll)

add_custom_command(TARGET ${PROJECT_NAME} POST_BUILD
                   COMMAND ${CMAKE_COMMAND} -E copy
                           ${CMAKE_BINARY_DIR}/libeible.dll
                           ${CMAKE_BINARY_DIR}/tests/libeible.dll)

install(TARGETS ${PROJECT_NAME}
    LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
    PUBLIC_HEADER DESTINATION ${CMAKE_INSTALL_INCLUDEDIR})

add_subdirectory(examples)

enable_testing()
add_subdirectory(tests)


