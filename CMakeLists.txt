cmake_minimum_required(VERSION 3.10)
project(FastExcel)

# 设置C++标准
set(CMAKE_CXX_STANDARD 11)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# 禁用第三方库的测试功能
set(SPDLOG_BUILD_TESTS OFF CACHE BOOL "" FORCE)
set(SPDLOG_BUILD_TESTS_HO OFF CACHE BOOL "" FORCE)
set(EXPAT_BUILD_TESTS OFF CACHE BOOL "" FORCE)
set(ZLIB_ENABLE_TESTS OFF CACHE BOOL "" FORCE)

add_subdirectory(third_party/spdlog)
add_subdirectory(third_party/libexpat/expat)
add_subdirectory(third_party/zlib-ng)

# 配置minizip-ng使用本地的zlib-ng
set(MZ_FETCH_LIBS OFF CACHE BOOL "" FORCE)
set(MZ_ZLIB ON CACHE BOOL "" FORCE)
set(MZ_FORCE_FETCH_LIBS OFF CACHE BOOL "" FORCE)
set(MZ_BUILD_TESTS OFF CACHE BOOL "" FORCE)
set(MZ_BUILD_UNIT_TESTS OFF CACHE BOOL "" FORCE)
set(MZ_BUILD_FUZZ_TESTS OFF CACHE BOOL "" FORCE)

# 设置zlib-ng库信息，让minizip-ng能找到它
set(ZLIB-NG_FOUND TRUE)
set(ZLIB-NG_LIBRARIES zlib)
set(ZLIB-NG_INCLUDE_DIRS ${CMAKE_CURRENT_SOURCE_DIR}/third_party/zlib-ng ${CMAKE_CURRENT_BINARY_DIR}/third_party/zlib-ng)
set(ZLIB-NG_LIBRARY_DIRS ${CMAKE_CURRENT_BINARY_DIR}/third_party/zlib-ng)

add_subdirectory(third_party/minizip-ng)

# 查找所有源文件
file(GLOB_RECURSE SOURCES "src/*.cpp" "src/*.c")

# 创建可执行文件
add_executable(${PROJECT_NAME} ${SOURCES})

# 链接third_party库
target_link_libraries(${PROJECT_NAME}
    PRIVATE
        spdlog::spdlog
        expat::expat
        zlib
        MINIZIP::minizip
)
