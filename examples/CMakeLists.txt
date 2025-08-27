# Examples CMakeLists.txt for FastExcel

# 收集所有示例源文件
# 精确列出保留的示例，避免构建过时/重复示例
set(EXAMPLE_SOURCES
        core_features_demo.cpp
        edit_features_demo.cpp
        formatting_features_demo.cpp
        image_insertion_demo.cpp
        performance_demo.cpp
        simd_test_demo.cpp
        memory_pool_advanced_demo.cpp
        minimal_destructor_test.cpp
        large_file_reading_demo.cpp
        readonly_memory_optimization_demo.cpp
)

# 为每个示例创建可执行文件
foreach (SOURCE ${EXAMPLE_SOURCES})
    get_filename_component(EXEC_NAME ${SOURCE} NAME_WE)

    add_executable(${EXEC_NAME} ${SOURCE})

    # 链接FastExcel库
    target_link_libraries(${EXEC_NAME} PRIVATE fastexcel)

    # 确保示例程序能访问第三方库的头文件
    target_include_directories(${EXEC_NAME} PRIVATE
            ${CMAKE_CURRENT_SOURCE_DIR}/../third_party/fmt/include
            ${FASTEXCEL_THIRD_PARTY_INCLUDE_DIRS}
    )

    # 链接第三方库
    target_link_libraries(${EXEC_NAME} PRIVATE
            fmt::fmt
            ${FASTEXCEL_EXPAT_TARGET}
    )

    # 为minizip_parallel_test添加minizip-ng链接
    if (EXEC_NAME STREQUAL "minizip_parallel_test" AND FASTEXCEL_MINIZIP_TARGET AND NOT FASTEXCEL_MINIZIP_TARGET STREQUAL "")
        target_link_libraries(${EXEC_NAME} PRIVATE ${FASTEXCEL_MINIZIP_TARGET})
    endif ()

    # 为compression_benchmark添加libdeflate链接（如果可用）
    if (EXEC_NAME STREQUAL "compression_benchmark" AND FASTEXCEL_LIBDEFLATE_TARGET AND NOT FASTEXCEL_LIBDEFLATE_TARGET STREQUAL "")
        target_link_libraries(${EXEC_NAME} PRIVATE ${FASTEXCEL_LIBDEFLATE_TARGET})
    endif ()

    # 为示例程序添加编译选项
    if (MSVC)
        # 设置源文件和执行字符集为UTF-8，解决C4819编码警告
        target_compile_options(${EXEC_NAME} PRIVATE /utf-8)
        # 禁用一些不必要的警告
        target_compile_options(${EXEC_NAME} PRIVATE /wd4251 /wd4275)
    else ()
        target_compile_options(${EXEC_NAME} PRIVATE -Wall -Wextra)
    endif ()

    # 设置输出目录
    set_target_properties(${EXEC_NAME} PROPERTIES
            RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin/examples
    )

    # 添加到示例组
    set_target_properties(${EXEC_NAME} PROPERTIES FOLDER "Examples")
endforeach ()

# 添加一个总的示例目标
add_custom_target(examples
        DEPENDS ${EXAMPLE_EXECUTABLES}
        COMMENT "Build all examples"
)
