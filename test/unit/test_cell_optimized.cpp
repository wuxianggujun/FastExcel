#include <gtest/gtest.h>
#include "fastexcel/core/Cell.hpp"
#include "fastexcel/core/FormatDescriptor.hpp"
#include "fastexcel/core/StyleBuilder.hpp"
#include <memory>
#include <chrono>
#include <iostream>

using namespace fastexcel::core;

class CellOptimizedTest : public ::testing::Test {
protected:
    void SetUp() override {
        cell = std::make_unique<Cell>();
    }
    
    std::unique_ptr<Cell> cell;
};

// 测试基本功能
TEST_F(CellOptimizedTest, BasicFunctionality) {
    // 测试空单元格
    EXPECT_TRUE(cell->isEmpty());
    EXPECT_EQ(cell->getType(), CellType::Empty);
    
    // 测试数字
    cell->setValue(42.5);
    EXPECT_TRUE(cell->isNumber());
    EXPECT_EQ(cell->getValue<double>(), 42.5);
    
    // 测试布尔值
    cell->setValue(true);
    EXPECT_TRUE(cell->isBoolean());
    EXPECT_TRUE(cell->getValue<bool>());
    
    // 测试清空
    cell->clear();
    EXPECT_TRUE(cell->isEmpty());
}

// 测试短字符串内联存储优化
TEST_F(CellOptimizedTest, InlineStringOptimization) {
    // 短字符串应该内联存储
    std::string short_str = "Hello";
    cell->setValue(short_str);
    
    EXPECT_TRUE(cell->isString());
    EXPECT_EQ(cell->getValue<std::string>(), short_str);
    EXPECT_EQ(cell->getInternalType(), CellType::InlineString);
    
    // 验证内存使用（应该很小）
    size_t memory_usage = cell->getMemoryUsage();
    EXPECT_LT(memory_usage, 100);  // 应该小于100字节
}

// 测试长字符串存储
TEST_F(CellOptimizedTest, LongStringStorage) {
    // 长字符串应该使用ExtendedData
    std::string long_str(100, 'A');  // 100个'A'
    cell->setValue(long_str);
    
    EXPECT_TRUE(cell->isString());
    EXPECT_EQ(cell->getValue<std::string>(), long_str);
    EXPECT_EQ(cell->getType(), CellType::String);
    
    // 验证内存使用（应该包含ExtendedData）
    size_t memory_usage = cell->getMemoryUsage();
    EXPECT_GT(memory_usage, 100);  // 应该大于100字节
}

// 测试公式功能
TEST_F(CellOptimizedTest, FormulaFunctionality) {
    cell->setFormula("=A1+B1", 10.5);
    
    EXPECT_TRUE(cell->isFormula());
    EXPECT_EQ(cell->getFormula(), "=A1+B1");
    EXPECT_EQ(cell->getFormulaResult(), 10.5);
    EXPECT_EQ(cell->getValue<double>(), 10.5);  // 应该返回公式结果
}

// 测试格式设置
TEST_F(CellOptimizedTest, FormatHandling) {
    auto format = std::make_shared<FormatDescriptor>(StyleBuilder().bold().build());
    
    // 测试shared_ptr版本
    cell->setFormat(format);
    EXPECT_TRUE(cell->hasFormat());
    EXPECT_EQ(cell->getFormatDescriptor(), format);
    // 只测试shared_ptr版本，删除了原始指针相关的测试
}

// 测试超链接功能
TEST_F(CellOptimizedTest, HyperlinkFunctionality) {
    std::string url = "https://example.com";
    
    cell->setHyperlink(url);
    EXPECT_TRUE(cell->hasHyperlink());
    EXPECT_EQ(cell->getHyperlink(), url);
    
    // 清空超链接
    cell->setHyperlink("");
    EXPECT_FALSE(cell->hasHyperlink());
    EXPECT_EQ(cell->getHyperlink(), "");
}

// 测试移动语义
TEST_F(CellOptimizedTest, MoveSemantics) {
    // 设置一个复杂的单元格
    cell->setValue("Test String");
    cell->setFormula("=A1+B1", 42.0);
    cell->setHyperlink("https://example.com");
    
    // 移动构造
    Cell moved_cell = std::move(*cell);
    EXPECT_TRUE(moved_cell.isFormula());
    EXPECT_EQ(moved_cell.getFormula(), "=A1+B1");
    EXPECT_EQ(moved_cell.getFormulaResult(), 42.0);
    EXPECT_TRUE(moved_cell.hasHyperlink());
    
    // 原对象应该被清空
    EXPECT_TRUE(cell->isEmpty());
}

// 测试拷贝语义
TEST_F(CellOptimizedTest, CopySemantics) {
    // 设置一个复杂的单元格
    cell->setValue("Test String");
    cell->setHyperlink("https://example.com");
    
    // 拷贝构造
    Cell copied_cell(*cell);
    EXPECT_TRUE(copied_cell.isString());
    EXPECT_EQ(copied_cell.getValue<std::string>(), "Test String");
    EXPECT_TRUE(copied_cell.hasHyperlink());
    EXPECT_EQ(copied_cell.getHyperlink(), "https://example.com");
    
    // 原对象应该保持不变
    EXPECT_TRUE(cell->isString());
    EXPECT_EQ(cell->getValue<std::string>(), "Test String");
}

// 测试内存使用优化
TEST_F(CellOptimizedTest, MemoryUsageOptimization) {
    // 空单元格应该使用最少内存
    size_t empty_usage = cell->getMemoryUsage();
    
    // 数字单元格应该使用相同内存
    cell->setValue(42.0);
    size_t number_usage = cell->getMemoryUsage();
    EXPECT_EQ(empty_usage, number_usage);
    
    // 短字符串应该使用相同内存
    cell->setValue("Hi");
    size_t short_string_usage = cell->getMemoryUsage();
    EXPECT_EQ(empty_usage, short_string_usage);
    
    // 长字符串应该使用更多内存
    cell->setValue(std::string(100, 'A'));
    size_t long_string_usage = cell->getMemoryUsage();
    EXPECT_GT(long_string_usage, empty_usage);
    
    // 创建另一个单元格测试公式内存使用
    Cell formula_cell;
    formula_cell.setFormula("=SUM(A1:A10)", 100.0);
    size_t formula_usage = formula_cell.getMemoryUsage();
    
    // 公式单元格应该比空单元格使用更多内存
    EXPECT_GT(formula_usage, empty_usage);
    
    // 测试短字符串vs短公式的内存使用
    Cell short_string_cell;
    short_string_cell.setValue("Hello");
    size_t short_str_usage = short_string_cell.getMemoryUsage();
    
    Cell short_formula_cell;
    short_formula_cell.setFormula("=A1+B1", 10.0);
    size_t short_formula_usage = short_formula_cell.getMemoryUsage();
    
    // 短公式应该比短字符串使用更多内存（因为需要ExtendedData）
    EXPECT_GT(short_formula_usage, short_str_usage);
}

// 性能基准测试
TEST_F(CellOptimizedTest, PerformanceBenchmark) {
    const int iterations = 10000;
    
    // 测试创建和销毁性能
    auto start = std::chrono::high_resolution_clock::now();
    
    for (int i = 0; i < iterations; ++i) {
        Cell temp_cell;
        temp_cell.setValue(static_cast<double>(i));
    }
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    // 应该能在合理时间内完成
    EXPECT_LT(duration.count(), 10000);  // 少于10ms
    
    std::cout << "Created and destroyed " << iterations 
              << " cells in " << duration.count() << " microseconds" << std::endl;
}