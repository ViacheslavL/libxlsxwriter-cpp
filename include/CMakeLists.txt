install(FILES xlsxwriter.h DESTINATION include)

install(DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}/xlsxwriter
    DESTINATION include
    FILES_MATCHING PATTERN "*.h")
