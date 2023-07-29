project(tests)

set(SOURCES
    Tests.cpp
    ExportXlsxTest.cpp
    Utilities.cpp
    TestTableModel.cpp
    ExportDsvTest.cpp
    ImportXlsxTest.cpp
    ImportOdsTest.cpp
    ImportCommon.cpp
)
qt_add_resources(SOURCES testResources.qrc)

set(HEADERS
    ExportXlsxTest.h
    Utilities.h
    TestTableModel.h
    ExportDsvTest.h
    ImportXlsxTest.h
    ImportOdsTest.h
    ImportCommon.h
 )

add_executable(${PROJECT_NAME} ${SOURCES} ${HEADERS})

target_link_libraries(${PROJECT_NAME} PRIVATE Qt6::Core Qt6::Widgets Qt6::Test QuaZip::QuaZip eible)

add_test(NAME ${PROJECT_NAME} COMMAND ${PROJECT_NAME})


