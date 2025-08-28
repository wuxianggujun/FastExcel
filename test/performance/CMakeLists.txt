cmake_minimum_required(VERSION 3.15)

# 检测 Google Benchmark 可用性（由 third_party 导出或系统包）
set(BENCHMARK_AVAILABLE OFF)
if (TARGET benchmark::benchmark)
    set(BENCHMARK_AVAILABLE ON)
elseif(DEFINED FASTEXCEL_BENCHMARK_TARGET AND NOT FASTEXCEL_BENCHMARK_TARGET STREQUAL "")
    set(BENCHMARK_AVAILABLE ON)
else()
    find_package(benchmark QUIET)
    if(benchmark_FOUND)
        set(BENCHMARK_AVAILABLE ON)
    endif()
endif()

#[[if(BENCHMARK_AVAILABLE)
    message(STATUS "Found Google Benchmark library")
    
    # 创建性能基准测试可执行程序
    add_executable(fastexcel_benchmark 
        performance_benchmark.cpp
    )
    
    # 选择可用的 benchmark 目标进行链接
    if (TARGET benchmark::benchmark)
        set(_BENCH_LIB benchmark::benchmark)
    elseif(DEFINED FASTEXCEL_BENCHMARK_TARGET AND NOT FASTEXCEL_BENCHMARK_TARGET STREQUAL "")
        set(_BENCH_LIB ${FASTEXCEL_BENCHMARK_TARGET})
    else()
        # 回退：find_package(benchmark) 提供的普通目标
        set(_BENCH_LIB benchmark)
    endif()
    target_link_libraries(fastexcel_benchmark PRIVATE fastexcel ${_BENCH_LIB})
    
    # 包含头文件目录
    target_include_directories(fastexcel_benchmark 
        PRIVATE 
        ${CMAKE_SOURCE_DIR}/src
    )
    
    # Windows 特定设置
    if(WIN32)
        target_link_libraries(fastexcel_benchmark PRIVATE psapi)
    endif()
    
    # 设置 C++17 标准
    target_compile_features(fastexcel_benchmark PRIVATE cxx_std_17)
    
    # 添加编译选项
    if(MSVC)
        # 仅在 Release/RelWithDebInfo 下启用优化，避免与 /RTC1 冲突（Debug 默认开启 /RTC1）
        target_compile_options(fastexcel_benchmark PRIVATE
            /W4 /WX-
            /wd4251 /wd4275   # 忽略 DLL 导出警告
            /utf-8            # 支持 UTF-8 编码
            $<$<CONFIG:Release>:/O2>
            $<$<CONFIG:RelWithDebInfo>:/O2>
        )
    else()
        # 非 MSVC 下 -g 与 -O2 可共存，这里始终启用 -O2
        target_compile_options(fastexcel_benchmark PRIVATE
            -Wall -Wextra -Wpedantic
            -O2               # 优化性能测试
        )
    endif()
    
    # 设置输出目录
    set_target_properties(fastexcel_benchmark PROPERTIES
        RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin
    )
    
    # 安装规则
    install(TARGETS fastexcel_benchmark
        RUNTIME DESTINATION bin
    )
    
    message(STATUS "FastExcel benchmark target created successfully")
    
else()
    message(WARNING "Google Benchmark not found. Performance benchmarks will not be built.")
    message(STATUS "To install Google Benchmark:")
    message(STATUS "  vcpkg: vcpkg install benchmark")
    message(STATUS "  Ubuntu: sudo apt-get install libbenchmark-dev")
    message(STATUS "  macOS: brew install google-benchmark")
endif()]]
