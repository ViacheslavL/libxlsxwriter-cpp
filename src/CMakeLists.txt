set(xlsxwriter_srcs
    app.cpp
    chart.cpp
    content_types.cpp
    core.cpp
    custom.cpp
    drawing.cpp
    format.cpp
    hash_table.cpp
    packager.cpp
    relationships.cpp
    shared_strings.cpp
    styles.cpp
    theme.cpp
    utility.cpp
    workbook.cpp
    worksheet.cpp
    xmlwriter.cpp)

if (MSVC)
    list(APPEND xlsxwriter_srcs XlsxWriter.def)
endif()

include_directories(${CMAKE_SOURCE_DIR}/include)

add_library(xlsxwriter ${xlsxwriter_srcs})

target_link_libraries(xlsxwriter ${ZLIB_LIBRARIES} lxw_tmpfileplus lxw_minizip)

add_library(xlsxwriter_test STATIC ${xlsxwriter_srcs})
target_compile_definitions(xlsxwriter_test PRIVATE -DTESTING)



install(TARGETS xlsxwriter
    EXPORT libxlsxwriter
    INCLUDES DESTINATION include
    ARCHIVE DESTINATION lib
    LIBRARY DESTINATION lib
    RUNTIME DESTINATION lib)
