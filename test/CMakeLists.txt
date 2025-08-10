# FastExcel 测试 CMakeLists.txt

cmake_minimum_required(VERSION 3.10)

# 检查是否启用了测试
if(NOT FASTEXCEL_BUILD_TESTS AND NOT FASTEXCEL_BUILD_UNIT_TESTS)
    message(STATUS "FastExcel: Tests are disabled")
    return()
endif()

# 确保 GoogleTest 目标可用
if(NOT DEFINED FASTEXCEL_GTEST_TARGETS)
    message(FATAL_ERROR "FastExcel: GoogleTest targets not available. Make sure third_party/CMakeLists.txt is processed first.")
endif()

# 包含目录
include_directories(${CMAKE_SOURCE_DIR}/src)
include_directories(${CMAKE_SOURCE_DIR}/third_party/fmt/include)
include_directories(${CMAKE_SOURCE_DIR}/third_party/libexpat/expat/lib)

# 单元测试源文件
set(UNIT_TEST_SOURCES
    unit/test_main.cpp
    unit/test_cell.cpp
    unit/test_cell_optimized.cpp
    unit/test_format.cpp
    unit/test_worksheet.cpp
    unit/test_workbook.cpp
    unit/test_xml_writer.cpp
    unit/test_xml_streaming.cpp
    unit/test_excel_zip_compatibility.cpp
)

# 性能测试源文件
set(PERFORMANCE_TEST_SOURCES
    unit/PerformanceTestSuite.cpp
    unit/test_performance_suite.cpp  # 性能测试主程序
)

# 添加性能测试子目录
add_subdirectory(performance)

# 集成测试源文件
set(INTEGRATION_TEST_SOURCES
    integration/test_integration.cpp
)

# 创建单元测试可执行文件
if(FASTEXCEL_BUILD_UNIT_TESTS)
    add_executable(fastexcel_unit_tests ${UNIT_TEST_SOURCES})
    
    # 链接库
    target_link_libraries(fastexcel_unit_tests
        fastexcel
        ${FASTEXCEL_GTEST_TARGETS}
    )
    
    # 设置 C++ 标准
    set_target_properties(fastexcel_unit_tests PROPERTIES
        CXX_STANDARD 17
        CXX_STANDARD_REQUIRED ON
    )
    
    # 为 MSVC 添加 UTF-8 编译选项
    if(MSVC)
        target_compile_options(fastexcel_unit_tests PRIVATE /utf-8)
    endif()
    
    # 添加测试
    if(CMAKE_VERSION VERSION_GREATER_EQUAL "3.10")
        include(GoogleTest)
        gtest_discover_tests(fastexcel_unit_tests)
    else()
        add_test(NAME fastexcel_unit_tests COMMAND fastexcel_unit_tests)
    endif()
    
    message(STATUS "FastExcel: Unit tests enabled")
endif()

# 创建集成测试可执行文件
if(FASTEXCEL_BUILD_INTEGRATION_TESTS)
    add_executable(fastexcel_integration_tests ${INTEGRATION_TEST_SOURCES})
    
    # 链接库
    target_link_libraries(fastexcel_integration_tests
        fastexcel
        ${FASTEXCEL_GTEST_TARGETS}
    )
    
    # 设置 C++ 标准
    set_target_properties(fastexcel_integration_tests PROPERTIES
        CXX_STANDARD 17
        CXX_STANDARD_REQUIRED ON
    )
    
    # 为 MSVC 添加 UTF-8 编译选项
    if(MSVC)
        target_compile_options(fastexcel_integration_tests PRIVATE /utf-8)
    endif()
    
    # 添加测试
    if(CMAKE_VERSION VERSION_GREATER_EQUAL "3.10")
        include(GoogleTest)
        gtest_discover_tests(fastexcel_integration_tests)
    else()
        add_test(NAME fastexcel_integration_tests COMMAND fastexcel_integration_tests)
    endif()
    
    message(STATUS "FastExcel: Integration tests enabled")
endif()

# 创建组合测试可执行文件（包含所有测试）
if(FASTEXCEL_BUILD_TESTS AND FASTEXCEL_BUILD_UNIT_TESTS)
    set(ALL_TEST_SOURCES ${UNIT_TEST_SOURCES})
    if(FASTEXCEL_BUILD_INTEGRATION_TESTS)
        list(APPEND ALL_TEST_SOURCES ${INTEGRATION_TEST_SOURCES})
    endif()
    
    add_executable(fastexcel_all_tests ${ALL_TEST_SOURCES})
    
    # 链接库
    target_link_libraries(fastexcel_all_tests
        fastexcel
        ${FASTEXCEL_GTEST_TARGETS}
    )
    
    # 设置 C++ 标准
    set_target_properties(fastexcel_all_tests PROPERTIES
        CXX_STANDARD 17
        CXX_STANDARD_REQUIRED ON
    )
    
    # 为 MSVC 添加 UTF-8 编译选项
    if(MSVC)
        target_compile_options(fastexcel_all_tests PRIVATE /utf-8)
    endif()
    
    # 添加测试
    if(CMAKE_VERSION VERSION_GREATER_EQUAL "3.10")
        include(GoogleTest)
        gtest_discover_tests(fastexcel_all_tests)
    else()
        add_test(NAME fastexcel_all_tests COMMAND fastexcel_all_tests)
    endif()
    
    message(STATUS "FastExcel: All tests enabled")
endif()

# 添加自定义测试目标
if(TARGET fastexcel_unit_tests)
    add_custom_target(run_unit_tests
        COMMAND fastexcel_unit_tests
        DEPENDS fastexcel_unit_tests
        COMMENT "运行 FastExcel 单元测试"
    )
endif()

if(TARGET fastexcel_integration_tests)
    add_custom_target(run_integration_tests
        COMMAND fastexcel_integration_tests
        DEPENDS fastexcel_integration_tests
        COMMENT "运行 FastExcel 集成测试"
    )
endif()

if(TARGET fastexcel_all_tests)
    add_custom_target(run_all_tests
        COMMAND fastexcel_all_tests
        DEPENDS fastexcel_all_tests
        COMMENT "运行 FastExcel 所有测试"
    )
endif()
