SET(PROJECT_EXAMPLES ${PROJECT_NAME}-examples)

set(${PROJECT_EXAMPLES}_SOURCES
    example.cpp
)

qt_add_resources(${PROJECT_EXAMPLES}_SOURCES examplesResources.qrc)

add_executable(${PROJECT_EXAMPLES} ${${PROJECT_EXAMPLES}_SOURCES})

target_link_libraries(${PROJECT_EXAMPLES} PRIVATE eible Qt6::Core Qt6::Widgets QuaZip)

if (WIN32)
    add_custom_command(TARGET ${PROJECT_EXAMPLES} POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E copy_directory
        "${CMAKE_CURRENT_BINARY_DIR}/../zlib/install/bin" "$<TARGET_FILE_DIR:${PROJECT_EXAMPLES}>")
    add_custom_command(TARGET ${PROJECT_EXAMPLES} POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E copy_directory
        "$<TARGET_FILE_DIR:QuaZip>" "$<TARGET_FILE_DIR:${PROJECT_EXAMPLES}>")
endif()
