SET(PROJECT_TESTS ${PROJECT_NAME}-tests)

set(${PROJECT_TESTS}_SOURCES
    Tests.cpp
    ExportXlsxTest.h
    ExportXlsxTest.cpp
    Utilities.h
    Utilities.cpp
    TestTableModel.h
    TestTableModel.cpp
    ExportDsvTest.h
    ExportDsvTest.cpp
    ImportXlsxTest.h
    ImportXlsxTest.cpp
    ImportOdsTest.h
    ImportOdsTest.cpp
    ImportCommon.h
    ImportCommon.cpp
)
qt_add_resources(${PROJECT_TESTS}_SOURCES testResources.qrc)

add_executable(${PROJECT_TESTS} ${${PROJECT_TESTS}_SOURCES})

target_link_libraries(${PROJECT_TESTS} PRIVATE eible Qt6::Core Qt6::Widgets Qt6::Test QuaZip)

add_test(NAME ${PROJECT_TESTS} COMMAND ${PROJECT_TESTS})


