# Examples CMakeLists.txt for FastExcel

# 收集所有示例源文件
file(GLOB EXAMPLE_SOURCES
    "*.cpp"
)

# 为每个示例创建可执行文件
foreach(SOURCE ${EXAMPLE_SOURCES})
    get_filename_component(EXEC_NAME ${SOURCE} NAME_WE)
    
    add_executable(${EXEC_NAME} ${SOURCE})
    
    # 链接FastExcel库
    target_link_libraries(${EXEC_NAME} PRIVATE fastexcel)
    
    # 设置输出目录
    set_target_properties(${EXEC_NAME} PROPERTIES
        RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin/examples
    )
    
    # 添加到示例组
    set_target_properties(${EXEC_NAME} PROPERTIES FOLDER "Examples")
endforeach()

# 添加一个总的示例目标
add_custom_target(examples 
    DEPENDS ${EXAMPLE_EXECUTABLES}
    COMMENT "Build all examples"
)