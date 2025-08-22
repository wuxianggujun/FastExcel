#include <gtest/gtest.h>
#include "fastexcel/FastExcel.hpp"
#include <filesystem>
#include <fstream>
#include <chrono>

using namespace fastexcel::core;

class IntegrationTest : public ::testing::Test {
protected:
    void SetUp() override {
        fastexcel::initialize();
        test_filename = "integration_test.xlsx";
    }
    
    void TearDown() override {
        // 清理测试文件
        if (std::filesystem::exists(test_filename)) {
            std::filesystem::remove(test_filename);
        }
        fastexcel::cleanup();
    }
    
    std::string test_filename;
};

// 测试完整的Excel文件生成流程
TEST_F(IntegrationTest, CompleteWorkflow) {
    // 创建工作簿
    auto workbook = Workbook::create(test_filename);
    ASSERT_NE(workbook, nullptr);
    
    EXPECT_TRUE(workbook->open());
    
    // 设置文档属性
    workbook->setTitle("集成测试报表");
    workbook->setAuthor("FastExcel测试");
    workbook->setCompany("测试公司");
    workbook->setCustomProperty("测试版本", "1.0");
    
    // 创建工作表
    auto worksheet = workbook->addWorksheet("测试数据");
    ASSERT_NE(worksheet, nullptr);
    
    // 创建格式
    auto header_format = workbook->createFormat();
    header_format->setBold(true);
    header_format->setBackgroundColor(COLOR_BLUE);
    header_format->setFontColor(COLOR_WHITE);
    header_format->setHorizontalAlign(HorizontalAlign::Center);
    
    auto currency_format = workbook->createFormat();
    currency_format->setNumberFormat("¥#,##0.00");
    
    auto date_format = workbook->createFormat();
    date_format->setNumberFormat("yyyy-mm-dd");
    
    // 写入表头
    worksheet->writeString(0, 0, "产品名称", header_format);
    worksheet->writeString(0, 1, "销售日期", header_format);
    worksheet->writeString(0, 2, "数量", header_format);
    worksheet->writeString(0, 3, "单价", header_format);
    worksheet->writeString(0, 4, "总额", header_format);
    
    // 写入数据
    std::vector<std::tuple<std::string, std::tm, int, double, double>> data = {
        {"产品A", {0, 0, 124, 1, 0, 0}, 100, 50.0, 5000.0},
        {"产品B", {0, 0, 124, 2, 0, 0}, 80, 75.0, 6000.0},
        {"产品C", {0, 0, 124, 3, 0, 0}, 120, 60.0, 7200.0}
    };
    
    for (size_t i = 0; i < data.size(); ++i) {
        int row = static_cast<int>(i + 1);
        worksheet->writeString(row, 0, std::get<0>(data[i]));
        worksheet->writeDateTime(row, 1, std::get<1>(data[i]), date_format);
        worksheet->writeNumber(row, 2, std::get<2>(data[i]));
        worksheet->writeNumber(row, 3, std::get<3>(data[i]), currency_format);
        worksheet->writeNumber(row, 4, std::get<4>(data[i]), currency_format);
    }
    
    // 添加总计行
    int total_row = static_cast<int>(data.size() + 1);
    worksheet->writeString(total_row, 0, "总计", header_format);
    worksheet->writeFormula(total_row, 4, "SUM(E2:E4)", currency_format);
    
    // 设置列宽
    worksheet->setColumnWidth(0, 15);
    worksheet->setColumnWidth(1, 12);
    worksheet->setColumnWidth(2, 8);
    worksheet->setColumnWidth(3, 12);
    worksheet->setColumnWidth(4, 12);
    
    // 设置自动筛选
    worksheet->setAutoFilter(0, 0, total_row, 4);
    
    // 冻结窗格
    worksheet->freezePanes(1, 0);
    
    // 保存文件
    EXPECT_TRUE(workbook->save());
    
    // 验证文件存在且大小合理
    EXPECT_TRUE(std::filesystem::exists(test_filename));
    auto file_size = std::filesystem::file_size(test_filename);
    EXPECT_GT(file_size, 1000); // 至少1KB
    EXPECT_LT(file_size, 1000000); // 不超过1MB
}

// 测试多工作表场景
TEST_F(IntegrationTest, MultipleWorksheets) {
    auto workbook = Workbook::create(test_filename);
    workbook->open();
    
    // 创建多个工作表
    auto sheet1 = workbook->addWorksheet("销售数据");
    auto sheet2 = workbook->addWorksheet("库存数据");
    auto sheet3 = workbook->addWorksheet("财务数据");
    
    // 在每个工作表中写入不同的数据
    sheet1->writeString(0, 0, "销售报表");
    sheet1->writeNumber(1, 0, 1000.0);
    
    sheet2->writeString(0, 0, "库存报表");
    sheet2->writeNumber(1, 0, 500.0);
    
    sheet3->writeString(0, 0, "财务报表");
    sheet3->writeNumber(1, 0, 2000.0);
    
    EXPECT_TRUE(workbook->save());
    EXPECT_TRUE(std::filesystem::exists(test_filename));
}

// 测试大数据量处理
TEST_F(IntegrationTest, LargeDataSet) {
    auto workbook = Workbook::create(test_filename);
    workbook->open();
    
    // 启用常量内存模式
    workbook->setConstantMemoryMode(true);
    
    auto worksheet = workbook->addWorksheet("大数据测试");
    
    const int rows = 1000;
    const int cols = 10;
    
    auto start_time = std::chrono::high_resolution_clock::now();
    
    // 写入大量数据
    for (int row = 0; row < rows; ++row) {
        for (int col = 0; col < cols; ++col) {
            if (col == 0) {
                worksheet->writeString(row, col, "Row " + std::to_string(row));
            } else {
                worksheet->writeNumber(row, col, row * cols + col);
            }
        }
    }
    
    auto end_time = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time);
    
    // 性能要求：1000x10的数据应该在合理时间内完成
    EXPECT_LT(duration.count(), 5000); // 少于5秒
    
    EXPECT_TRUE(workbook->save());
    EXPECT_TRUE(std::filesystem::exists(test_filename));
    
    // 验证文件大小合理
    auto file_size = std::filesystem::file_size(test_filename);
    EXPECT_GT(file_size, 10000); // 至少10KB
}

// 测试复杂格式化
TEST_F(IntegrationTest, ComplexFormatting) {
    auto workbook = Workbook::create(test_filename);
    workbook->open();
    
    auto worksheet = workbook->addWorksheet("格式测试");
    
    // 创建各种格式
    auto title_format = workbook->createFormat();
    title_format->setFontSize(18);
    title_format->setBold(true);
    title_format->setHorizontalAlign(HorizontalAlign::Center);
    title_format->setBackgroundColor(0x4472C4);
    title_format->setFontColor(COLOR_WHITE);
    
    auto border_format = workbook->createFormat();
    border_format->setBorder(BorderStyle::Thin);
    border_format->setBorderColor(COLOR_BLACK);
    
    auto number_format = workbook->createFormat();
    number_format->setNumberFormat("#,##0.00");
    number_format->setBorder(BorderStyle::Thin);
    
    auto percent_format = workbook->createFormat();
    percent_format->setNumberFormat("0.00%");
    percent_format->setBorder(BorderStyle::Thin);
    
    // 合并单元格作为标题
    worksheet->mergeRange(0, 0, 0, 4, "格式化测试报表", title_format);
    
    // 写入各种格式的数据
    worksheet->writeString(2, 0, "项目", border_format);
    worksheet->writeString(2, 1, "数值", border_format);
    worksheet->writeString(2, 2, "百分比", border_format);
    worksheet->writeString(2, 3, "货币", border_format);
    worksheet->writeString(2, 4, "日期", border_format);
    
    for (int i = 0; i < 5; ++i) {
        int row = 3 + i;
        worksheet->writeString(row, 0, "项目 " + std::to_string(i + 1), border_format);
        worksheet->writeNumber(row, 1, (i + 1) * 100.5, number_format);
        worksheet->writeNumber(row, 2, (i + 1) * 0.1, percent_format);
        worksheet->writeNumber(row, 3, (i + 1) * 1000.0, number_format);
        
        std::tm date = {};
        date.tm_year = 124;
        date.tm_mon = i;
        date.tm_mday = 1;
        worksheet->writeDateTime(row, 4, date);
    }
    
    // 设置列宽
    worksheet->setColumnWidth(0, 12);
    worksheet->setColumnWidth(1, 10);
    worksheet->setColumnWidth(2, 10);
    worksheet->setColumnWidth(3, 12);
    worksheet->setColumnWidth(4, 12);
    
    EXPECT_TRUE(workbook->save());
    EXPECT_TRUE(std::filesystem::exists(test_filename));
}

// 测试超链接功能
TEST_F(IntegrationTest, Hyperlinks) {
    auto workbook = Workbook::create(test_filename);
    workbook->open();
    
    auto worksheet = workbook->addWorksheet("超链接测试");
    
    auto link_format = workbook->createFormat();
    link_format->setFontColor(COLOR_BLUE);
    link_format->setUnderline(UnderlineType::Single);
    
    // 添加各种超链接
    worksheet->writeUrl(0, 0, "https://www.google.com", "Google", link_format);
    worksheet->writeUrl(1, 0, "https://www.github.com", "GitHub", link_format);
    worksheet->writeUrl(2, 0, "mailto:test@example.com", "发送邮件", link_format);
    worksheet->writeUrl(3, 0, "https://www.example.com"); // 不指定显示文本
    
    EXPECT_TRUE(workbook->save());
    EXPECT_TRUE(std::filesystem::exists(test_filename));
}

// 测试工作表保护
TEST_F(IntegrationTest, WorksheetProtection) {
    auto workbook = Workbook::create(test_filename);
    workbook->open();
    
    auto worksheet = workbook->addWorksheet("保护测试");
    
    // 写入一些数据
    worksheet->writeString(0, 0, "受保护的数据");
    worksheet->writeNumber(1, 0, 123.45);
    
    // 保护工作表
    worksheet->protect("password123");
    
    EXPECT_TRUE(workbook->save());
    EXPECT_TRUE(std::filesystem::exists(test_filename));
}

// 测试批量数据写入
TEST_F(IntegrationTest, BatchDataWrite) {
    auto workbook = Workbook::create(test_filename);
    workbook->open();
    
    auto worksheet = workbook->addWorksheet("批量数据");
    
    // 准备批量数据
    std::vector<std::vector<std::string>> string_data = {
        {"姓名", "部门", "职位"},
        {"张三", "销售部", "销售经理"},
        {"李四", "技术部", "软件工程师"},
        {"王五", "财务部", "会计师"}
    };
    
    std::vector<std::vector<double>> number_data = {
        {1.1, 2.2, 3.3},
        {4.4, 5.5, 6.6},
        {7.7, 8.8, 9.9}
    };
    
    // 批量写入字符串数据
    worksheet->writeRange(0, 0, string_data);
    
    // 批量写入数字数据
    worksheet->writeRange(5, 0, number_data);
    
    EXPECT_TRUE(workbook->save());
    EXPECT_TRUE(std::filesystem::exists(test_filename));
}

// 测试错误恢复
TEST_F(IntegrationTest, ErrorRecovery) {
    auto workbook = Workbook::create(test_filename);
    workbook->open();
    
    auto worksheet = workbook->addWorksheet("错误测试");
    
    // 写入正常数据
    worksheet->writeString(0, 0, "正常数据");
    
    // 尝试一些可能出错的操作
    try {
        worksheet->writeString(-1, 0, "无效位置"); // 应该抛出异常
        FAIL() << "应该抛出异常";
    } catch (const std::exception&) {
        // 预期的异常
    }
    
    // 验证工作簿仍然可用
    worksheet->writeString(1, 0, "恢复后的数据");
    
    EXPECT_TRUE(workbook->save());
    EXPECT_TRUE(std::filesystem::exists(test_filename));
}

// 测试内存管理
TEST_F(IntegrationTest, MemoryManagement) {
    // 创建和销毁多个工作簿，测试内存泄漏
    for (int i = 0; i < 10; ++i) {
        std::string filename = "memory_test_" + std::to_string(i) + ".xlsx";
        
        {
            auto workbook = Workbook::create(filename);
            workbook->open();
            
            auto worksheet = workbook->addWorksheet("测试");
            
            // 创建大量格式和数据
            for (int j = 0; j < 100; ++j) {
                auto format = workbook->createFormat();
                format->setBold(j % 2 == 0);
                worksheet->writeString(j, 0, "数据 " + std::to_string(j), format);
            }
            
            workbook->save();
        } // workbook应该在这里被自动销毁
        
        // 清理测试文件
        if (std::filesystem::exists(filename)) {
            std::filesystem::remove(filename);
        }
    }
    
    // 如果没有内存泄漏，这个测试应该正常完成
    SUCCEED();
}

// 测试并发访问（基本测试）
TEST_F(IntegrationTest, ConcurrentAccess) {
    std::vector<std::thread> threads;
    std::atomic<int> success_count(0);
    
    for (int i = 0; i < 3; ++i) {
        threads.emplace_back([i, &success_count]() {
            try {
                fastexcel::initialize();
                
                std::string filename = "concurrent_test_" + std::to_string(i) + ".xlsx";
                auto workbook = Workbook::create(filename);
                workbook->open();
                
                auto worksheet = workbook->addWorksheet("线程" + std::to_string(i));
                worksheet->writeString(0, 0, "线程 " + std::to_string(i) + " 的数据");
                
                if (workbook->save()) {
                    success_count++;
                }
                
                // 清理
                if (std::filesystem::exists(filename)) {
                    std::filesystem::remove(filename);
                }
                
                fastexcel::cleanup();
            } catch (...) {
                // 捕获任何异常
            }
        });
    }
    
    for (auto& thread : threads) {
        thread.join();
    }
    
    // 所有线程都应该成功
    EXPECT_EQ(success_count.load(), 3);
}

// 测试性能基准
TEST_F(IntegrationTest, PerformanceBenchmark) {
    auto workbook = Workbook::create(test_filename);
    workbook->open();
    workbook->setConstantMemoryMode(true);
    
    auto worksheet = workbook->addWorksheet("性能测试");
    
    const int test_rows = 5000;
    const int test_cols = 5;
    
    auto start_time = std::chrono::high_resolution_clock::now();
    
    // 写入测试数据
    for (int row = 0; row < test_rows; ++row) {
        worksheet->writeString(row, 0, "产品 " + std::to_string(row));
        worksheet->writeNumber(row, 1, row * 1.5);
        worksheet->writeNumber(row, 2, row * 2.0);
        worksheet->writeNumber(row, 3, row * 0.5);
        worksheet->writeFormula(row, 4, "B" + std::to_string(row + 1) + "*C" + std::to_string(row + 1));
    }
    
    auto write_time = std::chrono::high_resolution_clock::now();
    
    EXPECT_TRUE(workbook->save());
    
    auto save_time = std::chrono::high_resolution_clock::now();
    
    auto write_duration = std::chrono::duration_cast<std::chrono::milliseconds>(write_time - start_time);
    auto save_duration = std::chrono::duration_cast<std::chrono::milliseconds>(save_time - write_time);
    auto total_duration = std::chrono::duration_cast<std::chrono::milliseconds>(save_time - start_time);
    
    // 性能基准
    EXPECT_LT(write_duration.count(), 3000); // 写入应该少于3秒
    EXPECT_LT(save_duration.count(), 2000);  // 保存应该少于2秒
    EXPECT_LT(total_duration.count(), 5000); // 总时间应该少于5秒
    
    // 验证文件大小合理
    auto file_size = std::filesystem::file_size(test_filename);
    EXPECT_GT(file_size, 50000); // 至少50KB
    EXPECT_LT(file_size, 5000000); // 不超过5MB
    
    std::cout << "性能基准结果:" << std::endl;
    std::cout << "  写入时间: " << write_duration.count() << " 毫秒" << std::endl;
    std::cout << "  保存时间: " << save_duration.count() << " 毫秒" << std::endl;
    std::cout << "  总时间: " << total_duration.count() << " 毫秒" << std::endl;
    std::cout << "  文件大小: " << file_size << " 字节" << std::endl;
    std::cout << "  数据量: " << test_rows << " 行 x " << test_cols << " 列" << std::endl;
}