#include <gtest/gtest.h>
#include <chrono>
#include <iostream>

// 性能基准测试基础类
class PerformanceBenchmark : public ::testing::Test {
protected:
    void SetUp() override {
        start_time_ = std::chrono::high_resolution_clock::now();
    }

    void TearDown() override {
        auto end_time = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time_);
        
        std::cout << "⏱️  " << ::testing::UnitTest::GetInstance()->current_test_info()->name() 
                  << " 执行时间: " << duration.count() << " 微秒" << std::endl;
    }

    std::chrono::high_resolution_clock::time_point start_time_;
};

int main(int argc, char** argv) {
    std::cout << "🚀 FastExcel 性能基准测试套件" << std::endl;
    std::cout << "=====================================\n" << std::endl;
    
    ::testing::InitGoogleTest(&argc, argv);
    
    // 设置测试过滤器，默认运行所有基准测试
    if (argc == 1) {
        ::testing::GTEST_FLAG(filter) = "*Benchmark*";
    }
    
    int result = RUN_ALL_TESTS();
    
    std::cout << "\n🎉 性能基准测试完成！" << std::endl;
    return result;
}