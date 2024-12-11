SET(PROJECT_EXAMPLES ${PROJECT_NAME}-examples)

set(${PROJECT_EXAMPLES}_SOURCES
    example.cpp
)

qt_add_resources(${PROJECT_EXAMPLES}_SOURCES examplesResources.qrc)

add_executable(${PROJECT_EXAMPLES} ${${PROJECT_EXAMPLES}_SOURCES})

target_link_libraries(${PROJECT_EXAMPLES} PRIVATE eible Qt6::Core Qt6::Widgets QuaZip)
