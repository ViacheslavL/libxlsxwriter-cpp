add_library(lxw_ioapi ioapi.c)
add_library(lxw_zip zip.c)

include_directories(
    ${CMAKE_CURRENT_SOURCE_DIR}
    ${ZLIB_INCLUDE_DIR})

install(TARGETS lxw_ioapi lxw_zip
    INCLUDES DESTINATION include
    ARCHIVE DESTINATION lib
    LIBRARY DESTINATION lib)
