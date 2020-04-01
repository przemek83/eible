project(tests)

set(${PROJECT_NAME}_SOURCES
    Tests.cpp
    ExportXlsxTest.cpp
    ExportXlsxTest.h
    Utilities.cpp
    Utilities.h
    )

add_executable(${PROJECT_NAME} ${${PROJECT_NAME}_SOURCES})

target_include_directories(${PROJECT_NAME} PRIVATE ${CMAKE_CURRENT_SOURCE_DIR})
target_link_libraries(${PROJECT_NAME} eible Qt5::Test Qt5::Core Qt5::Widgets)
target_link_libraries(${PROJECT_NAME} optimized quazip5)
target_link_libraries(${PROJECT_NAME} debug quazip5d)

add_test(NAME ${PROJECT_NAME} COMMAND ${PROJECT_NAME})

