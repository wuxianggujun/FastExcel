# Third-party libraries configuration
cmake_minimum_required(VERSION 3.10)

# 构建选项
option(FASTEXCEL_USE_SYSTEM_LIBS "Use system libraries when available" OFF)

# 设置第三方库的构建类型
# 注意：所有第三方库强制构建为静态库以避免符号冲突和导入库问题
set(BUILD_SHARED_LIBS OFF CACHE BOOL "" FORCE)
message(STATUS "FastExcel: Building third-party libraries as static")

# =============================================================================
# fmt 配置
# =============================================================================
set(FMT_DOC OFF CACHE BOOL "" FORCE)
set(FMT_TEST OFF CACHE BOOL "" FORCE)
set(FMT_FUZZ OFF CACHE BOOL "" FORCE)
set(FMT_CUDA_TEST OFF CACHE BOOL "" FORCE)
set(FMT_OS ON CACHE BOOL "" FORCE)
set(FMT_MODULE OFF CACHE BOOL "" FORCE)
set(FMT_SYSTEM_HEADERS OFF CACHE BOOL "" FORCE)

if(FASTEXCEL_USE_SYSTEM_LIBS)
    find_package(fmt QUIET)
endif()

if(NOT fmt_FOUND)
    add_subdirectory(fmt)
    # 使用直接的目标名称
    set(FMT_TARGET fmt::fmt)
else()
    set(FMT_TARGET fmt::fmt)
    message(STATUS "FastExcel: Using system fmt")
endif()

# =============================================================================
# libexpat 配置
# =============================================================================
set(EXPAT_BUILD_TESTS OFF CACHE BOOL "" FORCE)
set(EXPAT_BUILD_TOOLS OFF CACHE BOOL "" FORCE)
set(EXPAT_BUILD_EXAMPLES OFF CACHE BOOL "" FORCE)
set(EXPAT_BUILD_DOCS OFF CACHE BOOL "" FORCE)
set(EXPAT_BUILD_FUZZERS OFF CACHE BOOL "" FORCE)

# 强制 expat 构建为静态库
set(EXPAT_BUILD_SHARED OFF CACHE BOOL "" FORCE)
set(EXPAT_SHARED_LIBS OFF CACHE BOOL "" FORCE)

if(FASTEXCEL_USE_SYSTEM_LIBS)
    find_package(expat QUIET)
endif()

if(NOT expat_FOUND)
    add_subdirectory(libexpat/expat)
    # 使用直接的目标名称
    set(EXPAT_TARGET expat)
else()
    set(EXPAT_TARGET expat::expat)
    message(STATUS "FastExcel: Using system expat")
endif()

# =============================================================================
# zlib 配置 (通过 minizip-ng 自动处理)
# =============================================================================
# minizip-ng 会自动处理 zlib 依赖，我们不需要单独配置 zlib-ng
if(FASTEXCEL_USE_SYSTEM_LIBS)
    find_package(ZLIB QUIET)
    if(ZLIB_FOUND)
        set(ZLIB_TARGET ZLIB::ZLIB)
        message(STATUS "FastExcel: Using system zlib")
    endif()
endif()

# 如果没有找到系统 zlib，minizip-ng 会自动获取
if(NOT ZLIB_FOUND)
    set(ZLIB_TARGET "")  # minizip-ng 会处理
endif()

# =============================================================================
# minizip-ng 配置
# =============================================================================
set(MZ_FETCH_LIBS OFF CACHE BOOL "" FORCE)  # 禁用自动获取，避免网络问题
set(MZ_ZLIB OFF CACHE BOOL "" FORCE)        # 禁用zlib依赖
set(MZ_BZIP2 OFF CACHE BOOL "" FORCE)       # 禁用bzip2
set(MZ_LZMA OFF CACHE BOOL "" FORCE)        # 禁用lzma
set(MZ_ZSTD OFF CACHE BOOL "" FORCE)        # 禁用zstd
set(MZ_LIBCOMP OFF CACHE BOOL "" FORCE)     # 禁用libcomp
set(MZ_BUILD_TESTS OFF CACHE BOOL "" FORCE)
set(MZ_BUILD_UNIT_TESTS OFF CACHE BOOL "" FORCE)
set(MZ_BUILD_FUZZ_TESTS OFF CACHE BOOL "" FORCE)
set(MZ_COMPAT ON CACHE BOOL "" FORCE)       # 启用兼容层
set(MZ_PKCRYPT ON CACHE BOOL "" FORCE)      # 启用传统加密，确保有基本功能
set(MZ_WZAES OFF CACHE BOOL "" FORCE)       # 禁用WinZIP AES加密
set(MZ_OPENSSL OFF CACHE BOOL "" FORCE)     # 禁用OpenSSL
set(MZ_COMPRESS_ONLY OFF CACHE BOOL "" FORCE)
set(MZ_DECOMPRESS_ONLY OFF CACHE BOOL "" FORCE)

add_subdirectory(minizip-ng)

# 确定正确的目标名称
set(MINIZIP_TARGET minizip)  # 由于 MZ_COMPAT=ON 且 MZ_LIB_SUFFIX=""，目标名称应该是 minizip

message(STATUS "FastExcel: minizip target set to: ${MINIZIP_TARGET}")

# =============================================================================
# GoogleTest 配置 (仅在需要测试时)
# =============================================================================
if(ENABLE_TESTS OR ENABLE_UNIT_TESTS OR FASTEXCEL_BUILD_TESTS)
    set(gtest_build_tests OFF CACHE BOOL "" FORCE)
    set(gtest_build_samples OFF CACHE BOOL "" FORCE)
    set(gtest_disable_pthreads ON CACHE BOOL "" FORCE)
    set(gtest_force_shared_crt ON CACHE BOOL "" FORCE)
    set(BUILD_GMOCK OFF CACHE BOOL "" FORCE)
    set(INSTALL_GTEST OFF CACHE BOOL "" FORCE)
    
    if(FASTEXCEL_USE_SYSTEM_LIBS)
        find_package(GTest QUIET)
    endif()
    
    if(NOT GTest_FOUND)
        add_subdirectory(googletest)
        set(GTEST_TARGETS gtest gtest_main)
    else()
        set(GTEST_TARGETS GTest::gtest GTest::gtest_main)
        message(STATUS "FastExcel: Using system GoogleTest")
    endif()
endif()

# =============================================================================
# 导出目标变量给父项目使用
# =============================================================================
set(FASTEXCEL_FMT_TARGET ${FMT_TARGET} PARENT_SCOPE)
set(FASTEXCEL_EXPAT_TARGET ${EXPAT_TARGET} PARENT_SCOPE)
set(FASTEXCEL_ZLIB_TARGET ${ZLIB_TARGET} PARENT_SCOPE)
set(FASTEXCEL_MINIZIP_TARGET ${MINIZIP_TARGET} PARENT_SCOPE)

if(DEFINED GTEST_TARGETS)
    set(FASTEXCEL_GTEST_TARGETS ${GTEST_TARGETS} PARENT_SCOPE)
endif()

# 导出包含目录
set(FASTEXCEL_THIRD_PARTY_INCLUDE_DIRS
    ${CMAKE_CURRENT_SOURCE_DIR}/libexpat/expat/lib
    ${CMAKE_CURRENT_SOURCE_DIR}/fmt/include
    ${CMAKE_CURRENT_SOURCE_DIR}/minizip-ng
    PARENT_SCOPE
)

message(STATUS "FastExcel: Third-party libraries configuration completed")