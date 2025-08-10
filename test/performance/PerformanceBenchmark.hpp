#pragma once
#include <gtest/gtest.h>
#include <chrono>
#include <iostream>

// 性能基准测试基类
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

private:
    std::chrono::high_resolution_clock::time_point start_time_;
};