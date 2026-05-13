cmake_minimum_required(VERSION 3.20)
project(comctl32v6hook LANGUAGES CXX)

if(NOT WIN32)
    message(FATAL_ERROR "This project targets Windows only.")
endif()

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# Dependency: Microsoft Detours.
#
# Recommended setup:
#   vcpkg install detours:arm64-windows detours:x64-windows detours:x86-windows
#   cmake -B build/arm64 -A ARM64 -DCMAKE_TOOLCHAIN_FILE=C:/vcpkg/scripts/buildsystems/vcpkg.cmake
find_path(DETOURS_INCLUDE_DIRS "detours/detours.h" REQUIRED)
find_library(DETOURS_LIBRARY_RELEASE detours REQUIRED
    PATHS "${_VCPKG_INSTALLED_DIR}/${VCPKG_TARGET_TRIPLET}/lib"
    NO_DEFAULT_PATH
)
find_library(DETOURS_LIBRARY_DEBUG detours REQUIRED
    PATHS "${_VCPKG_INSTALLED_DIR}/${VCPKG_TARGET_TRIPLET}/debug/lib"
    NO_DEFAULT_PATH
)

add_library(detours_imported STATIC IMPORTED)
set_target_properties(detours_imported PROPERTIES
    IMPORTED_LOCATION                 "${DETOURS_LIBRARY_RELEASE}"
    IMPORTED_LOCATION_DEBUG           "${DETOURS_LIBRARY_DEBUG}"
    IMPORTED_LOCATION_RELEASE         "${DETOURS_LIBRARY_RELEASE}"
    IMPORTED_LOCATION_RELWITHDEBINFO  "${DETOURS_LIBRARY_RELEASE}"
    IMPORTED_LOCATION_MINSIZEREL      "${DETOURS_LIBRARY_RELEASE}"
    INTERFACE_INCLUDE_DIRECTORIES     "${DETOURS_INCLUDE_DIRS}"
)

add_library(comctl32v6hook SHARED
    dllmain.cpp
)

# The DLL loads its activation context from a sibling .manifest file at runtime.
# Keep the manifest next to every build output so local testing and AppInit
# deployment use the same layout.
add_custom_command(TARGET comctl32v6hook POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E copy_if_different
        "${CMAKE_CURRENT_SOURCE_DIR}/comctl32v6hook.manifest"
        "$<TARGET_FILE_DIR:comctl32v6hook>/comctl32v6hook.manifest"
    VERBATIM
)

target_link_libraries(comctl32v6hook PRIVATE
    detours_imported
    user32
    kernel32
)

target_compile_options(comctl32v6hook PRIVATE
    /W4
    /WX
    /GS-
)

set_target_properties(comctl32v6hook PROPERTIES
    WINDOWS_EXPORT_ALL_SYMBOLS OFF
)

install(TARGETS comctl32v6hook
    RUNTIME DESTINATION bin
)
install(FILES comctl32v6hook.manifest
    DESTINATION bin
)
