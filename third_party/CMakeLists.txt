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
    set(SPDLOG_TARGET spdlog::spdlog)
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
    set(EXPAT_TARGET expat::expat)
else()
    set(EXPAT_TARGET expat::expat)
    message(STATUS "FastExcel: Using system expat")
endif()

# =============================================================================
# zlib-ng 配置
# =============================================================================
set(ZLIB_ENABLE_TESTS OFF CACHE BOOL "" FORCE)
set(ZLIBNG_ENABLE_TESTS OFF CACHE BOOL "" FORCE)
set(WITH_GTEST OFF CACHE BOOL "" FORCE)
set(WITH_FUZZERS OFF CACHE BOOL "" FORCE)
set(WITH_BENCHMARKS OFF CACHE BOOL "" FORCE)
set(WITH_BENCHMARK_APPS OFF CACHE BOOL "" FORCE)

if(FASTEXCEL_USE_SYSTEM_LIBS)
    find_package(ZLIB QUIET)
endif()

if(NOT ZLIB_FOUND)
    add_subdirectory(zlib-ng)
    set(ZLIB_TARGET zlib)
else()
    set(ZLIB_TARGET ZLIB::ZLIB)
    message(STATUS "FastExcel: Using system zlib")
endif()

# =============================================================================
# minizip-ng 配置
# =============================================================================
set(MZ_FETCH_LIBS OFF CACHE BOOL "" FORCE)
set(MZ_ZLIB ON CACHE BOOL "" FORCE)
set(MZ_FORCE_FETCH_LIBS OFF CACHE BOOL "" FORCE)
set(MZ_BUILD_TESTS OFF CACHE BOOL "" FORCE)
set(MZ_BUILD_UNIT_TESTS OFF CACHE BOOL "" FORCE)
set(MZ_BUILD_FUZZ_TESTS OFF CACHE BOOL "" FORCE)

# 设置zlib-ng库信息，让minizip-ng能找到它
set(ZLIB-NG_FOUND TRUE)
set(ZLIB-NG_LIBRARIES ${ZLIB_TARGET})
set(ZLIB-NG_INCLUDE_DIRS
    ${CMAKE_CURRENT_SOURCE_DIR}/zlib-ng
    ${CMAKE_CURRENT_BINARY_DIR}/zlib-ng
)
set(ZLIB-NG_LIBRARY_DIRS ${CMAKE_CURRENT_BINARY_DIR}/zlib-ng)

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
    ${CMAKE_CURRENT_SOURCE_DIR}/zlib-ng
    PARENT_SCOPE
)

message(STATUS "FastExcel: Third-party libraries configuration completed")