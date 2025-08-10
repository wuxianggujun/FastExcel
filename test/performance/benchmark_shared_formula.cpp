#include <gtest/gtest.h>
#include "PerformanceBenchmark.hpp"
#include "fastexcel/FastExcel.hpp"
#include "fastexcel/core/SharedFormula.hpp"
#include <chrono>
#include <memory>

class SharedFormulaBenchmark : public PerformanceBenchmark {
protected:
    void SetUp() override {
        PerformanceBenchmark::SetUp();
        workbook_ = fastexcel::core::Workbook::create(fastexcel::core::Path("benchmark_test.xlsx"));
        workbook_->open();
        worksheet_ = workbook_->addWorksheet("BenchmarkTest");
    }

    void TearDown() override {
        if (workbook_) {
            workbook_->save();
            workbook_->close();
        }
        PerformanceBenchmark::TearDown();
    }

    std::shared_ptr<fastexcel::core::Workbook> workbook_;
    std::shared_ptr<fastexcel::core::Worksheet> worksheet_;
};

// 测试共享公式创建性能
TEST_F(SharedFormulaBenchmark, CreateSharedFormulaPerformance) {
    const int FORMULA_COUNT = 1000;
    
    // 创建基础数据
    for (int i = 0; i < FORMULA_COUNT; ++i) {
        worksheet_->writeNumber(i, 0, i + 1);
        worksheet_->writeNumber(i, 1, (i + 1) * 2);
    }
    
    auto start = std::chrono::high_resolution_clock::now();
    
    // 创建共享公式 (1000行，公式: A1+B1, A2+B2, ...)
    worksheet_->createSharedFormula(0, 2, FORMULA_COUNT - 1, 2, "A1+B1");
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    std::cout << "📊 创建 " << FORMULA_COUNT << " 个共享公式单元格耗时: " 
              << duration.count() << " 微秒" << std::endl;
    
    // 验证创建成功
    EXPECT_TRUE(worksheet_->hasCellAt(0, 2));
    EXPECT_TRUE(worksheet_->hasCellAt(FORMULA_COUNT - 1, 2));
}

// 测试公式优化性能
TEST_F(SharedFormulaBenchmark, FormulaOptimizationPerformance) {
    const int FORMULA_COUNT = 500;
    
    // 添加基础数据
    for (int i = 0; i < FORMULA_COUNT; ++i) {
        worksheet_->writeNumber(i, 0, i + 1);
        worksheet_->writeNumber(i, 1, (i + 1) * 2);
    }
    
    // 添加相似的普通公式
    for (int i = 0; i < FORMULA_COUNT; ++i) {
        std::string formula = "A" + std::to_string(i + 1) + "+B" + std::to_string(i + 1);
        worksheet_->writeFormula(i, 2, formula);
    }
    
    auto start = std::chrono::high_resolution_clock::now();
    
    // 执行公式优化
    int optimized_count = worksheet_->optimizeFormulas(3);
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    std::cout << "⚡ 优化 " << FORMULA_COUNT << " 个公式耗时: " 
              << duration.count() << " 微秒，优化了 " << optimized_count << " 个公式" << std::endl;
    
    EXPECT_GT(optimized_count, 0);
}

// 测试大规模共享公式性能
TEST_F(SharedFormulaBenchmark, LargeSharedFormulaPerformance) {
    const int ROWS = 100;
    const int COLS = 100;
    const int TOTAL_CELLS = ROWS * COLS;
    
    // 添加基础数据（前两列）
    for (int i = 0; i < ROWS; ++i) {
        worksheet_->writeNumber(i, 0, i + 1);
        worksheet_->writeNumber(i, 1, (i + 1) * 2);
    }
    
    auto start = std::chrono::high_resolution_clock::now();
    
    // 创建大型共享公式 (100x100 = 10,000 cells)
    worksheet_->createSharedFormula(0, 2, ROWS - 1, COLS - 1, "A1+B1");
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);
    
    std::cout << "🚀 创建 " << TOTAL_CELLS << " 个共享公式单元格耗时: " 
              << duration.count() << " 毫秒" << std::endl;
    
    // 验证覆盖范围
    EXPECT_TRUE(worksheet_->hasCellAt(0, 2));
    EXPECT_TRUE(worksheet_->hasCellAt(ROWS - 1, COLS - 1));
}

// 测试公式模式检测性能
TEST_F(SharedFormulaBenchmark, PatternDetectionPerformance) {
    const int PATTERN_COUNT = 300;
    
    // 首先添加基础数据到D列和E列，避免循环引用
    for (int i = 0; i < PATTERN_COUNT; ++i) {
        worksheet_->writeNumber(i, 3, i + 1);        // D列：1, 2, 3, ...
        worksheet_->writeNumber(i, 4, (i + 1) * 2);  // E列：2, 4, 6, ...
    }
    
    // 创建三种不同的公式模式（避免自引用）
    for (int i = 0; i < PATTERN_COUNT; ++i) {
        // 模式1: 加法公式 (D1+E1, D2+E2, ...) - 引用有数据的列
        std::string formula1 = "D" + std::to_string(i + 1) + "+E" + std::to_string(i + 1);
        worksheet_->writeFormula(i, 0, formula1);
        
        // 模式2: 乘法公式 (D1*E1, D2*E2, ...) - 引用有数据的列
        std::string formula2 = "D" + std::to_string(i + 1) + "*E" + std::to_string(i + 1);
        worksheet_->writeFormula(i, 1, formula2);
        
        // 模式3: SUM函数公式 (SUM(D1:E1), SUM(D2:E2), ...) - 引用有数据的范围
        std::string formula3 = "SUM(D" + std::to_string(i + 1) + ":E" + std::to_string(i + 1) + ")";
        worksheet_->writeFormula(i, 2, formula3);
    }
    
    auto start = std::chrono::high_resolution_clock::now();
    
    // 分析优化潜力
    auto report = worksheet_->analyzeFormulaOptimization();
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    std::cout << "🔍 分析 " << PATTERN_COUNT * 3 << " 个公式的优化潜力耗时: " 
              << duration.count() << " 微秒" << std::endl;
    std::cout << "📈 发现 " << report.optimizable_formulas << " 个可优化公式，预估节省 " 
              << report.estimated_memory_savings << " 字节" << std::endl;
    
    EXPECT_GT(report.total_formulas, 0);
    EXPECT_GT(report.optimizable_formulas, 0);
}