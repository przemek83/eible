configure_file(CMakeLists.txt.zlib zlib/CMakeLists.txt)
execute_process(COMMAND ${CMAKE_COMMAND} -G ${CMAKE_GENERATOR} .
        WORKING_DIRECTORY ${CMAKE_CURRENT_BINARY_DIR}/zlib)
execute_process(COMMAND ${CMAKE_COMMAND} --build . --config Release
        WORKING_DIRECTORY ${CMAKE_CURRENT_BINARY_DIR}/zlib )

set(ZLIB_ROOT ${CMAKE_CURRENT_BINARY_DIR}/zlib/install)
include_directories( ${CMAKE_CURRENT_BINARY_DIR}/zlib/install/include)
