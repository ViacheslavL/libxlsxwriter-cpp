install(FILES xlsxwriter.hpp DESTINATION include)

install(DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}/xlsxwriter
    DESTINATION include
    FILES_MATCHING PATTERN "*.hpp")
