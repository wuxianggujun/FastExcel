cmake_minimum_required(VERSION 3.10)
project(FastExcel VERSION 1.0.0 LANGUAGES CXX C)

# =============================================================================
# 项目配置选项
# =============================================================================
option(FASTEXCEL_BUILD_SHARED_LIBS "Build FastExcel as shared library" OFF)
option(FASTEXCEL_USE_SYSTEM_LIBS "Use system libraries when available" OFF)
option(FASTEXCEL_USE_LIBDEFLATE "Enable libdeflate for high-performance compression" OFF)
option(FASTEXCEL_BUILD_EXAMPLES "Build example programs" ON)
option(FASTEXCEL_BUILD_TESTS "Build all tests" ON)
option(FASTEXCEL_BUILD_UNIT_TESTS "Build unit tests" ON)
option(FASTEXCEL_BUILD_INTEGRATION_TESTS "Build integration tests" OFF)
option(FASTEXCEL_INSTALL "Enable installation" OFF)

# =============================================================================
# 编译标准设置
# =============================================================================
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

set(CMAKE_C_STANDARD 11)
set(CMAKE_C_STANDARD_REQUIRED ON)

# =============================================================================
# 构建类型配置
# =============================================================================
if (NOT CMAKE_BUILD_TYPE)
    set(CMAKE_BUILD_TYPE "Release" CACHE STRING "Build type" FORCE)
endif ()

# 设置输出目录
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin)
set(CMAKE_LIBRARY_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/lib)
set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/lib)

# =============================================================================
# 第三方库配置
# =============================================================================
add_subdirectory(third_party)

# =============================================================================
# FastExcel 库源文件
# =============================================================================
set(FASTEXCEL_SOURCES
        src/fastexcel/FastExcel.cpp
        src/fastexcel/archive/ZipArchive.cpp
        src/fastexcel/archive/FileManager.cpp
        src/fastexcel/archive/MinizipParallelWriter.cpp
        src/fastexcel/archive/CompressionEngine.cpp
        src/fastexcel/archive/ZlibEngine.cpp
        src/fastexcel/core/Color.cpp
        src/fastexcel/core/Cell.cpp
        src/fastexcel/core/Exception.cpp
        src/fastexcel/core/Format.cpp
        src/fastexcel/core/Path.cpp
        src/fastexcel/core/Workbook.cpp
        src/fastexcel/core/Worksheet.cpp
        src/fastexcel/core/SharedStringTable.cpp
        src/fastexcel/core/FormatPool.cpp
        src/fastexcel/core/MemoryPool.cpp
        src/fastexcel/core/ErrorCode.cpp
        src/fastexcel/core/ThreadPool.cpp
        src/fastexcel/reader/XLSXReader.cpp
        src/fastexcel/reader/StylesParser.cpp
        src/fastexcel/reader/SharedStringsParser.cpp
        src/fastexcel/reader/WorksheetParser.cpp
        src/fastexcel/utils/Logger.cpp
        src/fastexcel/xml/ContentTypes.cpp
        src/fastexcel/xml/Relationships.cpp
        src/fastexcel/xml/SharedStrings.cpp
        src/fastexcel/xml/XMLStreamWriter.cpp
        src/fastexcel/xml/XMLStreamReader.cpp
        
)

# 如果启用了 libdeflate，添加相关源文件
if(FASTEXCEL_USE_LIBDEFLATE)
    list(APPEND FASTEXCEL_SOURCES
        src/fastexcel/archive/LibDeflateEngine.cpp
    )
endif()

# =============================================================================
# 创建 FastExcel 库
# =============================================================================
if (FASTEXCEL_BUILD_SHARED_LIBS)
    add_library(fastexcel SHARED ${FASTEXCEL_SOURCES})
    target_compile_definitions(fastexcel PRIVATE FASTEXCEL_EXPORTS)
    target_compile_definitions(fastexcel INTERFACE FASTEXCEL_SHARED)
    message(STATUS "FastExcel: Building shared library")
else ()
    add_library(fastexcel STATIC ${FASTEXCEL_SOURCES})
    message(STATUS "FastExcel: Building static library")
endif ()

# 设置库版本
set_target_properties(fastexcel PROPERTIES
        VERSION ${PROJECT_VERSION}
        SOVERSION ${PROJECT_VERSION_MAJOR}
        OUTPUT_NAME "fastexcel"
)

# =============================================================================
# 包含目录配置
# =============================================================================
target_include_directories(fastexcel PUBLIC
        $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/src>
        $<INSTALL_INTERFACE:include>
)

target_include_directories(fastexcel PRIVATE
        ${FASTEXCEL_THIRD_PARTY_INCLUDE_DIRS}
)

# 确保第三方库的头文件能被找到
target_include_directories(fastexcel PRIVATE
        ${CMAKE_CURRENT_SOURCE_DIR}/third_party/libexpat/expat/lib
        ${CMAKE_CURRENT_SOURCE_DIR}/third_party/fmt/include
        ${CMAKE_CURRENT_SOURCE_DIR}/third_party/minizip-ng
)

# =============================================================================
# 链接库配置
# =============================================================================
target_link_libraries(fastexcel PRIVATE
        fmt::fmt
        ${FASTEXCEL_EXPAT_TARGET}
)

# 只有当 minizip 目标存在时才链接
if (FASTEXCEL_MINIZIP_TARGET AND NOT FASTEXCEL_MINIZIP_TARGET STREQUAL "")
    target_link_libraries(fastexcel PRIVATE ${FASTEXCEL_MINIZIP_TARGET})
    message(STATUS "FastExcel: Linking with minizip target: ${FASTEXCEL_MINIZIP_TARGET}")
else ()
    message(WARNING "FastExcel: minizip target not available, ZipArchive functionality will be limited")
endif ()

# 链接 libdeflate（如果可用）
if (FASTEXCEL_LIBDEFLATE_TARGET AND NOT FASTEXCEL_LIBDEFLATE_TARGET STREQUAL "")
    target_link_libraries(fastexcel PRIVATE ${FASTEXCEL_LIBDEFLATE_TARGET})
    message(STATUS "FastExcel: Linking with libdeflate target: ${FASTEXCEL_LIBDEFLATE_TARGET}")
endif ()

# =============================================================================
# 宏定义配置
# =============================================================================
# 如果启用了 libdeflate，定义相应的宏
if(FASTEXCEL_USE_LIBDEFLATE AND FASTEXCEL_HAS_LIBDEFLATE)
    target_compile_definitions(fastexcel PRIVATE FASTEXCEL_HAS_LIBDEFLATE=1)
    message(STATUS "FastExcel: FASTEXCEL_HAS_LIBDEFLATE macro defined")
endif()


# =============================================================================
# 编译器配置
# =============================================================================
if (MSVC)
    target_compile_options(fastexcel PRIVATE /W4)
    # 禁用一些不必要的警告
    target_compile_options(fastexcel PRIVATE /wd4251 /wd4275)
    # 设置源文件和执行字符集为UTF-8，解决C4819编码警告
    target_compile_options(fastexcel PRIVATE /utf-8)
else ()
    target_compile_options(fastexcel PRIVATE -Wall -Wextra -Wpedantic)
endif ()

# =============================================================================
# 主程序
# =============================================================================
if (EXISTS "${CMAKE_CURRENT_SOURCE_DIR}/src/main.cpp")
    add_executable(${PROJECT_NAME} src/main.cpp)
    target_link_libraries(${PROJECT_NAME} PRIVATE fastexcel)

    # 确保主程序也能访问第三方库的头文件
    target_include_directories(${PROJECT_NAME} PRIVATE
            ${CMAKE_CURRENT_SOURCE_DIR}/third_party/fmt/include
            ${FASTEXCEL_THIRD_PARTY_INCLUDE_DIRS}
    )

    # 为主程序添加UTF-8编译选项
    if (MSVC)
        target_compile_options(${PROJECT_NAME} PRIVATE /utf-8)
    endif ()

    set_target_properties(${PROJECT_NAME} PROPERTIES
            OUTPUT_NAME "fastexcel_demo"
    )
endif ()

# =============================================================================
# 示例程序
# =============================================================================
if (FASTEXCEL_BUILD_EXAMPLES)
    add_subdirectory(examples)
endif ()

# =============================================================================
# 测试
# =============================================================================
if (FASTEXCEL_BUILD_TESTS)
    enable_testing()
    add_subdirectory(test)
endif ()

# =============================================================================
# 安装配置（可选）
# =============================================================================
if (FASTEXCEL_INSTALL)
    include(GNUInstallDirs)

    # 安装库文件
    install(TARGETS fastexcel
            LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
            ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR}
            RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR}
    )

    # 安装头文件
    install(DIRECTORY src/fastexcel
            DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}
            FILES_MATCHING PATTERN "*.hpp"
    )
endif ()

# =============================================================================
# 打印配置信息
# =============================================================================
message(STATUS "")
message(STATUS "FastExcel Configuration Summary:")
message(STATUS "  Version: ${PROJECT_VERSION}")
message(STATUS "  Build type: ${CMAKE_BUILD_TYPE}")
message(STATUS "  Shared libraries: ${FASTEXCEL_BUILD_SHARED_LIBS}")
message(STATUS "  Use system libs: ${FASTEXCEL_USE_SYSTEM_LIBS}")
message(STATUS "  Use libdeflate: ${FASTEXCEL_USE_LIBDEFLATE}")
message(STATUS "  Build examples: ${FASTEXCEL_BUILD_EXAMPLES}")
message(STATUS "  Build tests: ${FASTEXCEL_BUILD_TESTS}")
message(STATUS "  Install prefix: ${CMAKE_INSTALL_PREFIX}")
message(STATUS "")
