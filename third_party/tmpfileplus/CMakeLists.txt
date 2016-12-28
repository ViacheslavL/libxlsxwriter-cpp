add_library(lxw_tmpfileplus OBJECT tmpfileplus.c)

set_property(GLOBAL APPEND
    PROPERTY lxw_third_party_objects
    $<TARGET_OBJECTS:lxw_tmpfileplus>)
