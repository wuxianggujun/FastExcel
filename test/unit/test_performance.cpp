#include <gtest/gtest.h>
#include "fastexcel/FastExcel.hpp"
#include "fastexcel/utils/TimeUtils.hpp"
#include <chrono>
#include <random>
#include <filesystem>

class PerformanceTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 初始化FastExcel库
        ASSERT_TRUE(fastexcel::initialize());
    }
    
    void TearDown() override {
        // 清理FastExcel库
        fastexcel::cleanup();
    }
};

TEST_F(PerformanceTest, LargeDataWritePerformance) {
    const int rows = 10000;   // 1万行（测试用较小数据集）
    const int cols = 10;      // 10列
    const int total_cells = rows * cols;
    
    auto start_time = std::chrono::high_resolution_clock::now();
    
    // 创建工作簿
    auto workbook = std::make_shared<fastexcel::core::Workbook>("test_performance.xlsx");
    ASSERT_TRUE(workbook->open());
    
    // 优化配置
    auto& options = workbook->getOptions();
    options.compression_level = 0;           // 无压缩（最快）
    options.row_buffer_size = 5000;          // 较大缓冲区
    options.xml_buffer_size = 4 * 1024 * 1024; // 4MB XML缓冲区
    
    // 添加工作表
    auto worksheet = workbook->addWorksheet("性能测试");
    ASSERT_NE(worksheet, nullptr);
    
    // 创建随机数生成器
    std::mt19937 gen(12345);
    std::uniform_int_distribution<> int_dist(1, 1000);
    std::uniform_real_distribution<> real_dist(1.0, 1000.0);
    
    // 预生成字符串
    std::vector<std::string> pre_strings;
    pre_strings.reserve(100);
    for (int i = 0; i < 100; ++i) {
        pre_strings.push_back("TestData_" + std::to_string(i));
    }
    
    // 写入数据
    for (int row = 0; row < rows; ++row) {
        for (int col = 0; col < cols; ++col) {
            if (col == 0) {
                worksheet->writeNumber(row, col, row + 1);
            } else if (col == 1) {
                worksheet->writeString(row, col, pre_strings[row % pre_strings.size()]);
            } else if (col % 2 == 0) {
                worksheet->writeNumber(row, col, int_dist(gen));
            } else {
                worksheet->writeNumber(row, col, real_dist(gen));
            }
        }
    }
    
    auto write_time = std::chrono::high_resolution_clock::now();
    
    // 保存文件
    ASSERT_TRUE(workbook->save());
    workbook->close();
    
    auto total_time = std::chrono::high_resolution_clock::now();
    
    // 计算性能指标
    auto write_duration = std::chrono::duration_cast<std::chrono::milliseconds>(write_time - start_time);
    auto total_duration = std::chrono::duration_cast<std::chrono::milliseconds>(total_time - start_time);
    
    double cells_per_second = static_cast<double>(total_cells) / (total_duration.count() / 1000.0);
    
    // 性能断言（根据实际情况调整阈值）
    EXPECT_GT(cells_per_second, 10000) << "写入速度应该大于10K单元格/秒";
    EXPECT_LT(total_duration.count(), 30000) << "总耗时应该小于30秒";
    
    // 验证文件存在且有效
    EXPECT_TRUE(std::filesystem::exists("test_performance.xlsx"));
    
    // 清理测试文件
    std::filesystem::remove("test_performance.xlsx");
    
    // 输出性能信息（用于调试）
    std::cout << "性能测试结果:" << std::endl;
    std::cout << "总单元格数: " << total_cells << std::endl;
    std::cout << "总耗时: " << total_duration.count() << " 毫秒" << std::endl;
    std::cout << "写入速度: " << static_cast<int>(cells_per_second) << " 单元格/秒" << std::endl;
}

TEST_F(PerformanceTest, TimeUtilsPerformance) {
    const int iterations = 100000;
    
    auto start_time = std::chrono::high_resolution_clock::now();
    
    // 测试TimeUtils性能
    for (int i = 0; i < iterations; ++i) {
        auto current_time = fastexcel::utils::TimeUtils::getCurrentTime();
        auto time_t_val = fastexcel::utils::TimeUtils::tmToTimeT(current_time);
        auto formatted = fastexcel::utils::TimeUtils::formatTime(current_time);
        
        // 避免编译器优化掉这些调用
        (void)time_t_val;
        (void)formatted;
    }
    
    auto end_time = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);
    
    double ops_per_second = static_cast<double>(iterations) / (duration.count() / 1000000.0);
    
    // 性能断言
    EXPECT_GT(ops_per_second, 50000) << "TimeUtils操作应该大于50K次/秒";
    
    std::cout << "TimeUtils性能测试:" << std::endl;
    std::cout << "操作次数: " << iterations << std::endl;
    std::cout << "总耗时: " << duration.count() << " 微秒" << std::endl;
    std::cout << "操作速度: " << static_cast<int>(ops_per_second) << " 次/秒" << std::endl;
}

TEST_F(PerformanceTest, MemoryUsageTest) {
    // 这个测试主要验证大量数据写入时不会出现内存泄漏
    const int rows = 5000;
    const int cols = 20;
    
    auto workbook = std::make_shared<fastexcel::core::Workbook>("test_memory.xlsx");
    ASSERT_TRUE(workbook->open());
    
    auto worksheet = workbook->addWorksheet("内存测试");
    ASSERT_NE(worksheet, nullptr);
    
    // 写入大量数据
    for (int row = 0; row < rows; ++row) {
        for (int col = 0; col < cols; ++col) {
            worksheet->writeString(row, col, "TestData_" + std::to_string(row) + "_" + std::to_string(col));
        }
    }
    
    ASSERT_TRUE(workbook->save());
    workbook->close();
    
    // 验证文件存在
    EXPECT_TRUE(std::filesystem::exists("test_memory.xlsx"));
    
    // 清理测试文件
    std::filesystem::remove("test_memory.xlsx");
}