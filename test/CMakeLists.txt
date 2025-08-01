# Test CMakeLists.txt for FastExcel

# 启用测试
enable_testing()

# 包含GoogleTest
include(GoogleTest)

# 单元测试
if(FASTEXCEL_BUILD_UNIT_TESTS)
    # 收集单元测试源文件
    file(GLOB_RECURSE UNIT_TEST_SOURCES
        "unit/*.cpp"
    )
    
    # 创建单元测试可执行文件
    if(UNIT_TEST_SOURCES)
        add_executable(unit_tests ${UNIT_TEST_SOURCES})
        target_link_libraries(unit_tests PRIVATE
            fastexcel
            ${FASTEXCEL_GTEST_TARGETS}
        )
        target_include_directories(unit_tests PRIVATE
            ${CMAKE_SOURCE_DIR}/src
        )
        
        # 设置输出目录
        set_target_properties(unit_tests PROPERTIES
            RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin/test
            FOLDER "Tests"
        )
        
        # 设置运行时库为动态链接Debug版本
        if(MSVC)
            set_target_properties(unit_tests PROPERTIES
                MSVC_RUNTIME_LIBRARY "MultiThreadedDebugDLL"
            )
        endif()
        
        # 发现并注册测试
        gtest_discover_tests(unit_tests)
    endif()
endif()

# 集成测试
if(ENABLE_INTEGRATION_TESTS)
    # 收集集成测试源文件
    file(GLOB_RECURSE INTEGRATION_TEST_SOURCES
        "integration/*.cpp"
    )
    
    # 创建集成测试可执行文件
    if(INTEGRATION_TEST_SOURCES)
        add_executable(integration_tests ${INTEGRATION_TEST_SOURCES})
        target_link_libraries(integration_tests PRIVATE
            fastexcel
            ${FASTEXCEL_GTEST_TARGETS}
        )
        target_include_directories(integration_tests PRIVATE
            ${CMAKE_SOURCE_DIR}/src
        )
        
        # 设置输出目录
        set_target_properties(integration_tests PROPERTIES
            RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin/test
            FOLDER "Tests"
        )
        
        # 设置运行时库为动态链接Debug版本
        if(MSVC)
            set_target_properties(integration_tests PROPERTIES
                MSVC_RUNTIME_LIBRARY "MultiThreadedDebugDLL"
            )
        endif()
        
        # 发现并注册测试
        gtest_discover_tests(integration_tests)
    endif()
endif()


# 添加一个总的测试目标
add_custom_target(tests 
    DEPENDS unit_tests integration_tests
    COMMENT "Build all tests"
)

# 添加运行测试的目标
add_custom_target(run_tests
    COMMAND ${CMAKE_CTEST_COMMAND} --verbose
    DEPENDS tests
    COMMENT "Run all tests"
)
