# FastExcel 性能测试 CMakeLists.txt

cmake_minimum_required(VERSION 3.10)

# 检查是否启用了测试
if(NOT FASTEXCEL_BUILD_TESTS)
    message(STATUS "FastExcel: Performance tests are disabled")
    return()
endif()

# 确保 GoogleTest 目标可用
if(NOT DEFINED FASTEXCEL_GTEST_TARGETS)
    message(FATAL_ERROR "FastExcel: GoogleTest targets not available for performance tests.")
endif()

# 使用与test/CMakeLists.txt相同的全局包含目录设置方式
include_directories(${CMAKE_SOURCE_DIR}/src)
include_directories(${CMAKE_SOURCE_DIR}/third_party/fmt/include)
include_directories(${CMAKE_SOURCE_DIR}/third_party/libexpat/expat/lib)
include_directories(${CMAKE_SOURCE_DIR}/third_party/utfcpp/source)

# 性能测试源文件 - 完整版本
set(PERFORMANCE_TEST_SOURCES
    benchmark_shared_formula.cpp
    benchmark_xml_generation.cpp
    benchmark_read_write.cpp
    benchmark_main.cpp
)

# 创建性能测试可执行文件
add_executable(fastexcel_performance_tests ${PERFORMANCE_TEST_SOURCES})

# 链接库 - 只链接fastexcel库，所有依赖通过它传递
target_link_libraries(fastexcel_performance_tests
    fastexcel
    ${FASTEXCEL_GTEST_TARGETS}
)

# 设置 C++ 标准
set_target_properties(fastexcel_performance_tests PROPERTIES
    CXX_STANDARD 17
    CXX_STANDARD_REQUIRED ON
)

# 为 MSVC 添加 UTF-8 编译选项和宏定义
if(MSVC)
    target_compile_options(fastexcel_performance_tests PRIVATE /utf-8)
    target_compile_definitions(fastexcel_performance_tests PRIVATE 
        _CRT_SECURE_NO_WARNINGS
        NOMINMAX
    )
endif()

# 添加性能测试目标
add_custom_target(run_performance_tests
    COMMAND fastexcel_performance_tests --gtest_filter="*Benchmark*"
    DEPENDS fastexcel_performance_tests
    COMMENT "运行 FastExcel 性能基准测试"
)

message(STATUS "FastExcel: Performance benchmark tests enabled")