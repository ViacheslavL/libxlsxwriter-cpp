include_directories(
    ${CMAKE_CURRENT_SOURCE_DIR}
    ${ZLIB_INCLUDE_DIR})

set(minizip_sources
    ioapi.c
    zip.c)

if (MSVC)
    list(APPEND minizip_sources iowin32.c)
endif()

add_library(lxw_minizip STATIC ${minizip_sources})
