# Third-party libraries configuration
cmake_minimum_required(VERSION 3.10)

# 构建选项
option(FASTEXCEL_BUILD_SHARED_LIBS "Build shared libraries" ON)
option(FASTEXCEL_USE_SYSTEM_LIBS "Use system libraries when available" OFF)

# 设置第三方库的构建类型
if(FASTEXCEL_BUILD_SHARED_LIBS)
    set(BUILD_SHARED_LIBS ON CACHE BOOL "" FORCE)
    message(STATUS "FastExcel: Building with shared libraries")
else()
    set(BUILD_SHARED_LIBS OFF CACHE BOOL "" FORCE)
    message(STATUS "FastExcel: Building with static libraries")
endif()

# =============================================================================
# spdlog 配置
# =============================================================================
set(SPDLOG_BUILD_TESTS OFF CACHE BOOL "" FORCE)
set(SPDLOG_BUILD_TESTS_HO OFF CACHE BOOL "" FORCE)
set(SPDLOG_BUILD_EXAMPLE OFF CACHE BOOL "" FORCE)
set(SPDLOG_BUILD_BENCH OFF CACHE BOOL "" FORCE)

if(FASTEXCEL_USE_SYSTEM_LIBS)
    find_package(spdlog QUIET)
endif()

if(NOT spdlog_FOUND)
    add_subdirectory(spdlog)
    # 使用直接的目标名称而不是别名
    set(SPDLOG_TARGET spdlog)
else()
    set(SPDLOG_TARGET spdlog::spdlog)
    message(STATUS "FastExcel: Using system spdlog")
endif()

# =============================================================================
# libexpat 配置
# =============================================================================
set(EXPAT_BUILD_TESTS OFF CACHE BOOL "" FORCE)
set(EXPAT_BUILD_TOOLS OFF CACHE BOOL "" FORCE)
set(EXPAT_BUILD_EXAMPLES OFF CACHE BOOL "" FORCE)
set(EXPAT_BUILD_DOCS OFF CACHE BOOL "" FORCE)
set(EXPAT_BUILD_FUZZERS OFF CACHE BOOL "" FORCE)

# 根据构建类型设置 expat
if(FASTEXCEL_BUILD_SHARED_LIBS)
    set(EXPAT_BUILD_SHARED ON CACHE BOOL "" FORCE)
    set(EXPAT_SHARED_LIBS ON CACHE BOOL "" FORCE)
else()
    set(EXPAT_BUILD_SHARED OFF CACHE BOOL "" FORCE)
    set(EXPAT_SHARED_LIBS OFF CACHE BOOL "" FORCE)
endif()

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

add_subdirectory(minizip-ng)
set(MINIZIP_TARGET MINIZIP::minizip)

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
set(FASTEXCEL_SPDLOG_TARGET ${SPDLOG_TARGET} PARENT_SCOPE)
set(FASTEXCEL_EXPAT_TARGET ${EXPAT_TARGET} PARENT_SCOPE)
set(FASTEXCEL_ZLIB_TARGET ${ZLIB_TARGET} PARENT_SCOPE)
set(FASTEXCEL_MINIZIP_TARGET ${MINIZIP_TARGET} PARENT_SCOPE)

if(DEFINED GTEST_TARGETS)
    set(FASTEXCEL_GTEST_TARGETS ${GTEST_TARGETS} PARENT_SCOPE)
endif()

# 导出包含目录
set(FASTEXCEL_THIRD_PARTY_INCLUDE_DIRS
    ${CMAKE_CURRENT_SOURCE_DIR}/libexpat/expat/lib
    ${CMAKE_CURRENT_SOURCE_DIR}/spdlog/include
    ${CMAKE_CURRENT_SOURCE_DIR}/minizip-ng
    PARENT_SCOPE
)

message(STATUS "FastExcel: Third-party libraries configuration completed")