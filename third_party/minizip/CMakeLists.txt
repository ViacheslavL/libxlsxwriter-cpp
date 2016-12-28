include_directories(
    ${CMAKE_CURRENT_SOURCE_DIR}
    ${ZLIB_INCLUDE_DIR})

set(minizip_sources
    ioapi.c
    zip.c)

add_library(lxw_minizip OBJECT ${minizip_sources})

set_property(GLOBAL APPEND
    PROPERTY lxw_third_party_objects
    $<TARGET_OBJECTS:lxw_minizip>)
