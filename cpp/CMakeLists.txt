cmake_minimum_required(VERSION 3.20)
project(FF16SubsOrganizer CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

if(MSVC)
    add_compile_options(/W3 /utf-8 /EHsc)
    add_definitions(-D_CRT_SECURE_NO_WARNINGS -DNOMINMAX -DWIN32_LEAN_AND_MEAN)
else()
    add_compile_options(-Wall -Wextra)
endif()

include(FetchContent)
set(FETCHCONTENT_QUIET OFF)

# --- nlohmann/json (JSON parsing, header-only) ---
FetchContent_Declare(
    nlohmann_json
    GIT_REPOSITORY https://github.com/nlohmann/json.git
    GIT_TAG v3.11.3
    GIT_SHALLOW TRUE
)

# --- pugixml (XML parsing) ---
FetchContent_Declare(
    pugixml
    GIT_REPOSITORY https://github.com/zeux/pugixml.git
    GIT_TAG v1.14
    GIT_SHALLOW TRUE
)

# --- CLI11 (argparse, header-only) ---
FetchContent_Declare(
    CLI11
    GIT_REPOSITORY https://github.com/CLIUtils/CLI11.git
    GIT_TAG v2.4.2
    GIT_SHALLOW TRUE
)

# --- OpenXLSX (xlsx read/write) ---
set(OPENXLSX_CREATE_DOCS OFF CACHE BOOL "" FORCE)
set(OPENXLSX_BUILD_SAMPLES OFF CACHE BOOL "" FORCE)
set(OPENXLSX_BUILD_TESTS OFF CACHE BOOL "" FORCE)
set(OPENXLSX_BUILD_BENCHMARKS OFF CACHE BOOL "" FORCE)
FetchContent_Declare(
    OpenXLSX
    GIT_REPOSITORY https://github.com/troldal/OpenXLSX.git
    GIT_TAG v0.4.1
    GIT_SHALLOW TRUE
)

FetchContent_MakeAvailable(nlohmann_json pugixml CLI11 OpenXLSX)

add_executable(FF16SubsOrganizer
    src/main.cpp
    src/commands.cpp
    src/xml_io.cpp
    src/ids.cpp
    src/utils.cpp
)

target_include_directories(FF16SubsOrganizer PRIVATE src)

target_link_libraries(FF16SubsOrganizer PRIVATE
    nlohmann_json::nlohmann_json
    pugixml::pugixml
    CLI11::CLI11
    OpenXLSX::OpenXLSX
)

if(WIN32)
    target_compile_definitions(FF16SubsOrganizer PRIVATE
        _WIN32_WINNT=0x0A00
    )
endif()
