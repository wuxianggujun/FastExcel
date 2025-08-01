cmake_minimum_required(VERSION 3.10)
project(FastExcel)

# 设置C++标准
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# 禁用第三方库的测试功能
set(SPDLOG_BUILD_TESTS OFF CACHE BOOL "" FORCE)
set(SPDLOG_BUILD_TESTS_HO OFF CACHE BOOL "" FORCE)
set(EXPAT_BUILD_TESTS OFF CACHE BOOL "" FORCE)
set(ZLIB_ENABLE_TESTS OFF CACHE BOOL "" FORCE)

# 禁用 Gtest 自身测试
set(gtest_build_tests OFF CACHE BOOL "" FORCE)
set(gtest_build_samples OFF CACHE BOOL "" FORCE)
set(gtest_disable_pthreads ON CACHE BOOL "" FORCE)

# 强制 Gtest 使用共享运行时库
set(gtest_force_shared_crt ON CACHE BOOL "" FORCE)

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

# 添加 Gtest 支持
add_subdirectory(third_party/googletest)

# FastExcel库源文件
set(FASTEXCEL_SOURCES
    src/fastexcel/FastExcel.cpp
    src/fastexcel/archive/ZipArchive.cpp
    src/fastexcel/archive/FileManager.cpp
    src/fastexcel/core/Cell.cpp
    src/fastexcel/core/Format.cpp
    src/fastexcel/core/Workbook.cpp
    src/fastexcel/core/Worksheet.cpp
    src/fastexcel/utils/Logger.cpp
    src/fastexcel/xml/ContentTypes.cpp
    src/fastexcel/xml/Relationships.cpp
    src/fastexcel/xml/SharedStrings.cpp
    src/fastexcel/xml/XMLStreamWriter.cpp
)

# 创建静态库
add_library(fastexcel STATIC ${FASTEXCEL_SOURCES})

# 包含目录
target_include_directories(fastexcel PUBLIC
        ${CMAKE_CURRENT_SOURCE_DIR}/src
        ${CMAKE_CURRENT_SOURCE_DIR}/third_party/libexpat/expat/lib
        ${CMAKE_CURRENT_SOURCE_DIR}/third_party/spdlog/include
        ${CMAKE_CURRENT_SOURCE_DIR}/third_party/minizip-ng
        ${CMAKE_CURRENT_SOURCE_DIR}/third_party/zlib-ng
)

# 链接库
target_link_libraries(fastexcel PUBLIC
        spdlog::spdlog
        expat::expat
        zlib
        MINIZIP::minizip
)

# 添加编译器警告
if(MSVC)
    target_compile_options(fastexcel PRIVATE /W4)
else()
    target_compile_options(fastexcel PRIVATE -Wall -Wextra)
endif()

# 示例程序
add_executable(${PROJECT_NAME} src/main.cpp)
target_link_libraries(${PROJECT_NAME} PRIVATE fastexcel)

# 添加构建选项
option(ENABLE_EXAMPLES "Enable building of examples" OFF)
option(ENABLE_TESTS "Enable building of all tests" ON)
option(ENABLE_UNIT_TESTS "Enable building of unit tests" ON)
option(ENABLE_INTEGRATION_TESTS "Enable building of integration tests" OFF)

# 添加examples子目录
if (ENABLE_EXAMPLES)
    add_subdirectory(examples)
endif ()


# 如果总开关开启，则所有子开关都开启
if (ENABLE_TESTS)
    add_subdirectory(test)
endif ()

