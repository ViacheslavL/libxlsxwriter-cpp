add_library(lxw_tmpfileplus tmpfileplus.c)

install(TARGETS lxw_tmpfileplus
    INCLUDES DESTINATION include
    ARCHIVE DESTINATION lib
    LIBRARY DESTINATION lib)
