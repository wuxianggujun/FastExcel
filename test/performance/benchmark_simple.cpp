#include <gtest/gtest.h>
#include "PerformanceBenchmark.hpp"
#include <chrono>
#include <memory>
#include <vector>
#include <string>

// 简单的性能基准测试，避免使用可能有问题的头文件
class SimpleBenchmark : public PerformanceBenchmark {
protected:
    void SetUp() override {
        PerformanceBenchmark::SetUp();
    }

    void TearDown() override {
        PerformanceBenchmark::TearDown();
    }
};

// 测试字符串操作性能
TEST_F(SimpleBenchmark, StringOperationPerformance) {
    const int COUNT = 100000;
    
    auto start = std::chrono::high_resolution_clock::now();
    
    std::vector<std::string> strings;
    strings.reserve(COUNT);
    
    for (int i = 0; i < COUNT; ++i) {
        strings.push_back("Cell_" + std::to_string(i / 100) + "_" + std::to_string(i % 100));
    }
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    std::cout << "🚀 创建 " << COUNT << " 个字符串耗时: " 
              << duration.count() << " 微秒" << std::endl;
    
    EXPECT_EQ(strings.size(), COUNT);
    EXPECT_FALSE(strings[0].empty());
}

// 测试数值计算性能
TEST_F(SimpleBenchmark, NumericalCalculationPerformance) {
    const int ROWS = 1000;
    const int COLS = 100;
    
    auto start = std::chrono::high_resolution_clock::now();
    
    std::vector<std::vector<double>> matrix(ROWS, std::vector<double>(COLS));
    
    for (int row = 0; row < ROWS; ++row) {
        for (int col = 0; col < COLS; ++col) {
            matrix[row][col] = row * col + 1.5;
        }
    }
    
    // 简单计算
    double sum = 0.0;
    for (int row = 0; row < ROWS; ++row) {
        for (int col = 0; col < COLS; ++col) {
            sum += matrix[row][col];
        }
    }
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    std::cout << "🔢 矩阵计算 " << ROWS << "x" << COLS << " 耗时: " 
              << duration.count() << " 微秒，结果: " << sum << std::endl;
    
    EXPECT_GT(sum, 0.0);
}

// 测试内存分配性能
TEST_F(SimpleBenchmark, MemoryAllocationPerformance) {
    const int COUNT = 10000;
    
    auto start = std::chrono::high_resolution_clock::now();
    
    std::vector<std::unique_ptr<std::vector<int>>> containers;
    containers.reserve(COUNT);
    
    for (int i = 0; i < COUNT; ++i) {
        auto container = std::make_unique<std::vector<int>>();
        container->resize(100);
        for (int j = 0; j < 100; ++j) {
            (*container)[j] = i + j;
        }
        containers.push_back(std::move(container));
    }
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    std::cout << "💾 内存分配 " << COUNT << " 个容器耗时: " 
              << duration.count() << " 微秒" << std::endl;
    
    EXPECT_EQ(containers.size(), COUNT);
    EXPECT_EQ(containers[0]->size(), 100);
}