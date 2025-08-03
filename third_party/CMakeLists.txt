# Third-party libraries configuration
cmake_minimum_required(VERSION 3.10)

# 设置 CMake 策略以允许跨目录链接
cmake_policy(SET CMP0079 NEW)

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
# ZlibMinizipBundle 配置
# =============================================================================
# 配置ZlibMinizipBundle选项
set(ZLIBMINIZIPBUNDLE_BUILD_TESTS OFF CACHE BOOL "" FORCE)

add_subdirectory(ZlibMinizipBundle)

# 设置目标变量
set(ZLIB_TARGET zlibstatic)
set(MINIZIP_TARGET minizip)
set(ZLIBMINIZIPBUNDLE_TARGET ZlibMinizipBundle::ZlibMinizipBundle)

message(STATUS "FastExcel: Using ZlibMinizipBundle")
message(STATUS "  ZLIB target: ${ZLIB_TARGET}")
message(STATUS "  MINIZIP target: ${MINIZIP_TARGET}")
message(STATUS "  Bundle target: ${ZLIBMINIZIPBUNDLE_TARGET}")

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

# 如果使用ZlibMinizipBundle，也导出Bundle目标
if(DEFINED ZLIBMINIZIPBUNDLE_TARGET)
    set(FASTEXCEL_ZLIBMINIZIPBUNDLE_TARGET ${ZLIBMINIZIPBUNDLE_TARGET} PARENT_SCOPE)
endif()

if(DEFINED GTEST_TARGETS)
    set(FASTEXCEL_GTEST_TARGETS ${GTEST_TARGETS} PARENT_SCOPE)
endif()

# 导出包含目录
set(FASTEXCEL_THIRD_PARTY_INCLUDE_DIRS
    ${CMAKE_CURRENT_SOURCE_DIR}/libexpat/expat/lib
    ${CMAKE_CURRENT_SOURCE_DIR}/fmt/include
    ${CMAKE_CURRENT_SOURCE_DIR}/ZlibMinizipBundle/third_party/zlib-ng
    ${CMAKE_CURRENT_BINARY_DIR}/ZlibMinizipBundle/third_party/zlib-ng
    ${CMAKE_CURRENT_SOURCE_DIR}/ZlibMinizipBundle/third_party/minizip-ng
    PARENT_SCOPE
)

message(STATUS "FastExcel: Third-party libraries configuration completed")