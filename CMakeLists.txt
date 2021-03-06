cmake_minimum_required(VERSION 3.0.0)

project(xlsxwriter++)

set(CMAKE_DEBUG_POSTFIX d)

# force out of source build so we don't trash the existing Makefiles
if ("${CMAKE_BINARY_DIR}" STREQUAL "${CMAKE_SOURCE_DIR}")
    message(FATAL_ERROR " This project requires an out of source build."
        " Create an empty directory and configure and build from there.")
endif()

# if we forgot to set the build type default to release
if (NOT CMAKE_BUILD_TYPE)
    set(CMAKE_BUILD_TYPE "Release"
      CACHE STRING "One of: Debug Release RelWithDebInfo MinSizeRel"
      FORCE)
    set_property(CACHE CMAKE_BUILD_TYPE PROPERTY
        STRINGS "Debug" "Release" "MinSizeRel" "RelWithDebInfo")
endif()
message(STATUS "Configuring a ${CMAKE_BUILD_TYPE} build")

# locate dependencies
find_package(ZLIB REQUIRED)
include_directories(${ZLIB_INCLUDE_DIRS})

# for configuring shared libraries
set(CMAKE_MACOSX_RPATH 1)
set(BUILD_SHARED_LIBS ON CACHE BOOL "Build shared libraries")

# set up build dirs
set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/lib)
set(CMAKE_LIBRARY_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/lib)
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin)
set(CMAKE_EXPORT_ALL_SYMBOLS TRUE)

# set default compiler flags
if (NOT MSVC)
    set(tmp "-fPIC -Wall -Wextra -O0")
    if (NOT ((APPLE) AND ("${CMAKE_C_COMPILER_ID}" MATCHES "Clang"))
        AND ("${CMAKE_BUILD_TYPE}" MATCHES "Release"))
        set(tmp "${tmp} -mtune=native")
    endif()
    set(CMAKE_C_FLAGS "${tmp}"
        CACHE STRING "c-libraries flags compile"
        FORCE)
    set(CMAKE_CXX_FLAGS "${tmp} -std=c++0x"
        CACHE STRING "libxlsxwriter compile flags"
        FORCE)
else()
    set(CMAKE_C_FLAGS "/c /Zi /nologo /D _USRDLL /D _WINDLL /D UNICODE /D _UNICODE /WX- /D NDEBUG /D _CRT_SECURE_NO_WARNINGS /Gm- /EHsc /MD /GS /Gy- /fp:precise /Zc:wchar_t /Zc:forScope /Gz /TC")
    set(CMAKE_C_FLAGS_RELEASE "/O2 /Ot")
    set(CMAKE_CXX_FLAGS "/c /Zi /nologo /D _USRDLL /D _WINDLL /D UNICODE /D _UNICODE /WX- /D NDEBUG /D _CRT_SECURE_NO_WARNINGS /Gm- /EHsc /MD /GS /Gy- /fp:precise /Zc:wchar_t /Zc:forScope /Gz")
    set(CMAKE_CXX_FLAGS_RELEASE "/O2 /Ot")
    set(CMAKE_EXE_LINKER_FLAGS "/NOLOGO /DEBUG /SUBSYSTEM:WINDOWS /TLBID:1 /DYNAMICBASE /NXCOMPAT")
endif()

# for building tests and examples
function(add_simple_executable ex_name)
    string(CONCAT ex_source ${ex_name} ".cpp")
    add_executable(${ex_name} ${ex_source})
    target_link_libraries(${ex_name} xlsxwriter++ ${ZLIB_LIBRARIES})
endfunction()

# process library components
add_subdirectory(third_party)
add_subdirectory(src)
add_subdirectory(include)
#add_subdirectory(examples)

# enable regression tests
set(BUILD_TESTING OFF CACHE BOOL "Enable regresion tests")
if (BUILD_TESTING)
    #include(CTest)
    configure_file(
        "${CMAKE_CURRENT_SOURCE_DIR}/CTestCustom.cmake.in"
        "${CMAKE_CURRENT_BINARY_DIR}/CTestCustom.cmake"
        @ONLY)
    add_subdirectory(test)
endif()
