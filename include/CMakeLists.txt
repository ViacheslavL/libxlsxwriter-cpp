add_custom_target(headers SOURCES
    xlsxwriter.hpp
    xlsxwriter/app.hpp
    xlsxwriter/chart.hpp
    xlsxwriter/content_types.hpp
    xlsxwriter/core.hpp
    xlsxwriter/custom.hpp
    xlsxwriter/drawing.hpp
    xlsxwriter/format.hpp
    xlsxwriter/hash_table.hpp
    xlsxwriter/packager.hpp
    xlsxwriter/relationships.hpp
    xlsxwriter/shared_strings.hpp
    xlsxwriter/styles.hpp
    xlsxwriter/theme.hpp
    xlsxwriter/utility.hpp
    xlsxwriter/workbook.hpp
    xlsxwriter/worksheet.hpp
    xlsxwriter/xmlwriter.hpp)

install(FILES xlsxwriter.hpp DESTINATION include)

install(DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}/xlsxwriter
    DESTINATION include
    FILES_MATCHING PATTERN "*.hpp")


