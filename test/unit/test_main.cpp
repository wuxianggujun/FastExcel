#include <gtest/gtest.h>
#include <iostream>

// 测试主函数
int main(int argc, char** argv) {
    std::cout << "FastExcel 单元测试开始..." << std::endl;
    
    // 初始化 GoogleTest
    ::testing::InitGoogleTest(&argc, argv);
    
    // 运行所有测试
    int result = RUN_ALL_TESTS();
    
    if (result == 0) {
        std::cout << "所有测试通过！" << std::endl;
    } else {
        std::cout << "有测试失败！" << std::endl;
    }
    
    return result;
}