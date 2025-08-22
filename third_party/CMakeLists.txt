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
# libdeflate 配置
# =============================================================================
option(FASTEXCEL_USE_LIBDEFLATE "Enable libdeflate for high-performance compression" ON)

if(FASTEXCEL_USE_LIBDEFLATE)
    # 配置 libdeflate 选项
    set(LIBDEFLATE_BUILD_STATIC_LIB ON CACHE BOOL "" FORCE)
    set(LIBDEFLATE_BUILD_SHARED_LIB OFF CACHE BOOL "" FORCE)
    set(LIBDEFLATE_BUILD_GZIP OFF CACHE BOOL "" FORCE)
    set(LIBDEFLATE_BUILD_TESTS OFF CACHE BOOL "" FORCE)
    
    # 为libdeflate设置高性能编译选项（在add_subdirectory之前）
    if(MSVC)
        # 临时保存当前的CMAKE标志
        set(ORIGINAL_CMAKE_C_FLAGS_DEBUG "${CMAKE_C_FLAGS_DEBUG}")
        set(ORIGINAL_CMAKE_CXX_FLAGS_DEBUG "${CMAKE_CXX_FLAGS_DEBUG}")
        
        # 为libdeflate设置优化的编译标志
        set(CMAKE_C_FLAGS_DEBUG "/O2 /DNDEBUG /arch:AVX2")
        set(CMAKE_CXX_FLAGS_DEBUG "/O2 /DNDEBUG /arch:AVX2")
        
        message(STATUS "FastExcel: Using optimized MSVC flags for libdeflate")
    else()
        # 临时保存当前的CMAKE标志
        set(ORIGINAL_CMAKE_C_FLAGS_DEBUG "${CMAKE_C_FLAGS_DEBUG}")
        set(ORIGINAL_CMAKE_CXX_FLAGS_DEBUG "${CMAKE_CXX_FLAGS_DEBUG}")
        
        # 为libdeflate设置优化的编译标志
        set(CMAKE_C_FLAGS_DEBUG "-O3 -DNDEBUG -march=native")
        set(CMAKE_CXX_FLAGS_DEBUG "-O3 -DNDEBUG -march=native")
        
        message(STATUS "FastExcel: Using optimized GCC/Clang flags for libdeflate")
    endif()
    
    if(FASTEXCEL_USE_SYSTEM_LIBS)
        find_package(libdeflate QUIET)
    endif()
    
    if(NOT libdeflate_FOUND)
        add_subdirectory(libdeflate)
        set(LIBDEFLATE_TARGET libdeflate_static)
        message(STATUS "FastExcel: Using bundled libdeflate with high-performance optimizations")
        
        # 恢复原始的CMAKE标志，避免影响其他目标
        set(CMAKE_C_FLAGS_DEBUG "${ORIGINAL_CMAKE_C_FLAGS_DEBUG}")
        set(CMAKE_CXX_FLAGS_DEBUG "${ORIGINAL_CMAKE_CXX_FLAGS_DEBUG}")
        
    else()
        set(LIBDEFLATE_TARGET libdeflate::libdeflate)
        message(STATUS "FastExcel: Using system libdeflate")
        
        # 恢复原始的CMAKE标志
        set(CMAKE_C_FLAGS_DEBUG "${ORIGINAL_CMAKE_C_FLAGS_DEBUG}")
        set(CMAKE_CXX_FLAGS_DEBUG "${ORIGINAL_CMAKE_CXX_FLAGS_DEBUG}")
    endif()
    
    # 定义宏以启用 libdeflate 支持
    set(FASTEXCEL_HAS_LIBDEFLATE ON PARENT_SCOPE)
    message(STATUS "FastExcel: libdeflate support enabled")
else()
    message(STATUS "FastExcel: libdeflate support disabled")
endif()

# =============================================================================
# UTF8-CPP 配置 (header-only library)
# =============================================================================
# UTF8-CPP 是一个纯头文件库，用于处理UTF-8编码
set(UTF8CPP_BUILD_TESTS OFF CACHE BOOL "" FORCE)
set(UTF8CPP_BUILD_SAMPLES OFF CACHE BOOL "" FORCE)

if(FASTEXCEL_USE_SYSTEM_LIBS)
    find_package(utf8cpp QUIET)
endif()

if(NOT utf8cpp_FOUND)
    add_subdirectory(utfcpp)
    set(UTF8CPP_TARGET utf8cpp)
    message(STATUS "FastExcel: Using bundled UTF8-CPP")
else()
    set(UTF8CPP_TARGET utf8cpp::utf8cpp)
    message(STATUS "FastExcel: Using system UTF8-CPP")
endif()

# =============================================================================
# Highway SIMD Library 配置
# =============================================================================
option(FASTEXCEL_USE_HIGHWAY "Enable Highway SIMD library for high-performance operations" ON)

if(FASTEXCEL_USE_HIGHWAY)
    # 配置 Highway 选项
    set(HWY_ENABLE_TESTS OFF CACHE BOOL "" FORCE)
    set(HWY_ENABLE_EXAMPLES OFF CACHE BOOL "" FORCE)
    set(HWY_ENABLE_CONTRIB OFF CACHE BOOL "" FORCE)
    set(HWY_ENABLE_INSTALL OFF CACHE BOOL "" FORCE)
    
    if(FASTEXCEL_USE_SYSTEM_LIBS)
        find_package(hwy QUIET)
    endif()
    
    if(NOT hwy_FOUND)
        add_subdirectory(highway)
        set(HIGHWAY_TARGET hwy)
        message(STATUS "FastExcel: Using bundled Highway SIMD library")
    else()
        set(HIGHWAY_TARGET hwy::hwy)
        message(STATUS "FastExcel: Using system Highway SIMD library")
    endif()
    
    # 定义宏以启用 Highway 支持
    set(FASTEXCEL_HAS_HIGHWAY ON PARENT_SCOPE)
    message(STATUS "FastExcel: Highway SIMD support enabled")
else()
    message(STATUS "FastExcel: Highway SIMD support disabled")
endif()

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

# 导出 libdeflate 目标
if(DEFINED LIBDEFLATE_TARGET)
    set(FASTEXCEL_LIBDEFLATE_TARGET ${LIBDEFLATE_TARGET} PARENT_SCOPE)
endif()

# 导出 UTF8CPP 目标
if(DEFINED UTF8CPP_TARGET)
    set(FASTEXCEL_UTF8CPP_TARGET ${UTF8CPP_TARGET} PARENT_SCOPE)
endif()

# 导出 Highway 目标
if(DEFINED HIGHWAY_TARGET)
    set(FASTEXCEL_HIGHWAY_TARGET ${HIGHWAY_TARGET} PARENT_SCOPE)
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
    ${CMAKE_CURRENT_SOURCE_DIR}/utfcpp/source
    PARENT_SCOPE
)

# 如果启用了 libdeflate，添加其包含目录
if(FASTEXCEL_USE_LIBDEFLATE AND DEFINED LIBDEFLATE_TARGET)
    set(FASTEXCEL_THIRD_PARTY_INCLUDE_DIRS
        ${FASTEXCEL_THIRD_PARTY_INCLUDE_DIRS}
        ${CMAKE_CURRENT_SOURCE_DIR}/libdeflate
        PARENT_SCOPE
    )
endif()

# 如果启用了 Highway，添加其包含目录
if(FASTEXCEL_USE_HIGHWAY AND DEFINED HIGHWAY_TARGET)
    set(FASTEXCEL_THIRD_PARTY_INCLUDE_DIRS
        ${FASTEXCEL_THIRD_PARTY_INCLUDE_DIRS}
        ${CMAKE_CURRENT_SOURCE_DIR}/highway
        PARENT_SCOPE
    )
endif()

message(STATUS "FastExcel: Third-party libraries configuration completed")
