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

include_directories(${CMAKE_SOURCE_DIR}/include)

add_library(xlsxwriter++ ${xlsxwriter_srcs})

target_link_libraries(xlsxwriter++ lxw_tmpfileplus lxw_minizip ${ZLIB_LIBRARIES}  )

target_compile_definitions(xlsxwriter++ PRIVATE -DXLSXWRITER_EXPORTS )

add_library(xlsxwriter_test STATIC ${xlsxwriter_srcs})

target_compile_definitions(xlsxwriter_test PRIVATE -DTESTING -DXLSXWRITER_EXPORTS)



install(TARGETS xlsxwriter++
    EXPORT xlsxwriter++
    INCLUDES DESTINATION include
    ARCHIVE DESTINATION lib
    LIBRARY DESTINATION lib
    RUNTIME DESTINATION lib)
