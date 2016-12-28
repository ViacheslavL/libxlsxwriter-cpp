add_subdirectory(minizip)

set(USE_STANDARD_TMPFILE OFF CACHE BOOL
    "Enable/Disable the tmpfileplus third party library")

if (NOT USE_STANDARD_TMPFILE)
    add_subdirectory(tmpfileplus)
endif()
