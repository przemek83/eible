project(tests)

qt5_add_resources(RESOURCES_TESTS testResources.qrc)
add_custom_target(generatedTest_qrc DEPENDS ${RESOURCES_TESTS})

set(${PROJECT_NAME}_SOURCES
    qrc_testResources.cpp
    Tests.cpp
    ExportXlsxTest.cpp
    ExportXlsxTest.h
    Utilities.cpp
    Utilities.h
    TestTableModel.cpp
    TestTableModel.h
    ExportDsvTest.cpp
    ExportDsvTest.h
    ImportXlsxTest.cpp
    ImportXlsxTest.h
    ImportOdsTest.cpp
    ImportOdsTest.h
    ImportCommon.cpp
    ImportCommon.h
    )

add_executable(${PROJECT_NAME} ${${PROJECT_NAME}_SOURCES})

add_dependencies(${PROJECT_NAME} generatedTest_qrc)

target_include_directories(${PROJECT_NAME} PRIVATE ${CMAKE_CURRENT_SOURCE_DIR})
target_link_libraries(${PROJECT_NAME} eible Qt5::Test Qt5::Core Qt5::Widgets)
target_link_libraries(${PROJECT_NAME} optimized Qt5Quazip)
target_link_libraries(${PROJECT_NAME} debug Qt5Quazipd)

add_test(NAME ${PROJECT_NAME} COMMAND ${PROJECT_NAME})

