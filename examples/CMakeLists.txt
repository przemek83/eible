project(examples)

set(SOURCES
    example.cpp
)
qt_add_resources(SOURCES examplesResources.qrc)

add_executable(${PROJECT_NAME} ${SOURCES})

target_link_libraries(${PROJECT_NAME} PRIVATE Qt6::Core Qt6::Widgets QuaZip::QuaZip eible)
