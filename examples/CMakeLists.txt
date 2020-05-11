project(examples)

qt5_add_resources(RESOURCES_EXAMPLES examplesResources.qrc)
add_custom_target(generatedExamples_qrc DEPENDS ${RESOURCES_EXAMPLES})

set(${PROJECT_NAME}_SOURCES
    qrc_examplesResources.cpp
    example.cpp
    )

add_executable(examples ${${PROJECT_NAME}_SOURCES})

add_dependencies(${PROJECT_NAME} generatedExamples_qrc)

target_link_libraries(examples PRIVATE Qt5::Core Qt5::Widgets eible)
target_link_libraries(${PROJECT_NAME} PRIVATE optimized Qt5Quazip)
target_link_libraries(${PROJECT_NAME} PRIVATE debug Qt5Quazipd)
