/**
 * @file worksheet_optimization_example.cpp
 * @brief 展示Worksheet集成优化功能的示例
 * 
 * 本示例演示了如何使用集成到Worksheet中的优化特性：
 * 1. 优化的Cell类（75%内存减少）
 * 2. 共享字符串表（SST）去重
 * 3. 格式池去重
 * 4. 优化模式的使用
 * 5. 性能统计和监控
 */

#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/SharedStringTable.hpp"
#include "fastexcel/core/FormatPool.hpp"
#include "fastexcel/core/Format.hpp"
#include <iostream>
#include <chrono>
#include <vector>
#include <random>
#include <iomanip>
#include <memory>

using namespace fastexcel::core;

// 性能测试辅助类
class PerformanceTimer {
public:
    PerformanceTimer(const std::string& name) : name_(name) {
        start_ = std::chrono::high_resolution_clock::now();
    }
    
    ~PerformanceTimer() {
        auto end = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start_);
        std::cout << "[" << name_ << "] 耗时: " << duration.count() << "ms" << std::endl;
    }

private:
    std::string name_;
    std::chrono::high_resolution_clock::now_type::time_point start_;
};

// 内存使用格式化函数
std::string formatMemory(size_t bytes) {
    if (bytes < 1024) {
        return std::to_string(bytes) + " B";
    } else if (bytes < 1024 * 1024) {
        return std::to_string(bytes / 1024) + " KB";
    } else {
        return std::to_string(bytes / (1024 * 1024)) + " MB";
    }
}

// 演示基本优化功能
void demonstrateBasicOptimization() {
    std::cout << "\n========== 基本优化功能演示 ==========" << std::endl;
    
    // 创建工作簿和优化组件
    auto workbook = std::make_shared<Workbook>();
    SharedStringTable sst;
    FormatPool format_pool;
    
    // 创建工作表并设置优化组件
    auto worksheet = std::make_unique<Worksheet>("优化测试", workbook, 1);
    worksheet->setSharedStringTable(&sst);
    worksheet->setFormatPool(&format_pool);
    
    std::cout << "创建工作表，设置优化组件完成" << std::endl;
    
    // 创建一些格式
    auto header_format = std::make_shared<Format>();
    header_format->setBold(true);
    header_format->setFontColor(0xFFFFFF);
    header_format->setBackgroundColor(0x4472C4);
    
    auto data_format = std::make_shared<Format>();
    data_format->setFontSize(11);
    
    auto number_format = std::make_shared<Format>();
    number_format->setNumberFormat("0.00");
    
    // 写入测试数据
    {
        PerformanceTimer timer("写入测试数据");
        
        // 写入表头
        worksheet->writeString(0, 0, "产品名称", header_format);
        worksheet->writeString(0, 1, "价格", header_format);
        worksheet->writeString(0, 2, "数量", header_format);
        worksheet->writeString(0, 3, "总计", header_format);
        
        // 写入数据行
        std::vector<std::string> products = {"苹果", "香蕉", "橙子", "葡萄", "西瓜"};
        for (int i = 0; i < 1000; ++i) {
            int row = i + 1;
            std::string product = products[i % products.size()];
            double price = 2.0 + (i % 10) * 0.5;
            int quantity = 10 + (i % 20);
            
            worksheet->writeString(row, 0, product, data_format);
            worksheet->writeNumber(row, 1, price, number_format);
            worksheet->writeNumber(row, 2, quantity, data_format);
            worksheet->writeFormula(row, 3, "B" + std::to_string(row + 1) + "*C" + std::to_string(row + 1), number_format);
        }
    }
    
    // 获取性能统计
    auto stats = worksheet->getPerformanceStats();
    std::cout << "\n性能统计:" << std::endl;
    std::cout << "  总单元格数: " << stats.total_cells << std::endl;
    std::cout << "  内存使用: " << formatMemory(stats.memory_usage) << std::endl;
    std::cout << "  SST字符串数: " << stats.sst_strings << std::endl;
    std::cout << "  SST压缩比: " << std::fixed << std::setprecision(1) << stats.sst_compression_ratio << "%" << std::endl;
    std::cout << "  唯一格式数: " << stats.unique_formats << std::endl;
    std::cout << "  格式去重比: " << std::fixed << std::setprecision(1) << stats.format_deduplication_ratio << "%" << std::endl;
}

// 演示优化模式
void demonstrateOptimizeMode() {
    std::cout << "\n========== 优化模式演示 ==========" << std::endl;
    
    auto workbook = std::make_shared<Workbook>();
    SharedStringTable sst;
    FormatPool format_pool;
    
    // 创建两个工作表进行对比
    auto standard_sheet = std::make_unique<Worksheet>("标准模式", workbook, 1);
    auto optimized_sheet = std::make_unique<Worksheet>("优化模式", workbook, 2);
    
    // 设置优化组件
    standard_sheet->setSharedStringTable(&sst);
    standard_sheet->setFormatPool(&format_pool);
    optimized_sheet->setSharedStringTable(&sst);
    optimized_sheet->setFormatPool(&format_pool);
    
    // 启用优化模式
    optimized_sheet->setOptimizeMode(true);
    
    const int ROWS = 5000;
    const int COLS = 10;
    
    // 标准模式性能测试
    {
        PerformanceTimer timer("标准模式写入");
        for (int row = 0; row < ROWS; ++row) {
            for (int col = 0; col < COLS; ++col) {
                if (col % 2 == 0) {
                    standard_sheet->writeString(row, col, "数据_" + std::to_string(row % 100));
                } else {
                    standard_sheet->writeNumber(row, col, row * col * 0.01);
                }
            }
        }
    }
    
    // 优化模式性能测试
    {
        PerformanceTimer timer("优化模式写入");
        for (int row = 0; row < ROWS; ++row) {
            for (int col = 0; col < COLS; ++col) {
                if (col % 2 == 0) {
                    optimized_sheet->writeString(row, col, "数据_" + std::to_string(row % 100));
                } else {
                    optimized_sheet->writeNumber(row, col, row * col * 0.01);
                }
            }
        }
        // 刷新最后一行
        optimized_sheet->flushCurrentRow();
    }
    
    // 性能对比
    auto standard_stats = standard_sheet->getPerformanceStats();
    auto optimized_stats = optimized_sheet->getPerformanceStats();
    
    std::cout << "\n性能对比:" << std::endl;
    std::cout << "标准模式:" << std::endl;
    std::cout << "  单元格数: " << standard_stats.total_cells << std::endl;
    std::cout << "  内存使用: " << formatMemory(standard_stats.memory_usage) << std::endl;
    
    std::cout << "优化模式:" << std::endl;
    std::cout << "  单元格数: " << optimized_stats.total_cells << std::endl;
    std::cout << "  内存使用: " << formatMemory(optimized_stats.memory_usage) << std::endl;
    
    if (standard_stats.memory_usage > 0) {
        double memory_reduction = (1.0 - static_cast<double>(optimized_stats.memory_usage) / standard_stats.memory_usage) * 100.0;
        std::cout << "  内存减少: " << std::fixed << std::setprecision(1) << memory_reduction << "%" << std::endl;
    }
}

// 演示大数据量处理
void demonstrateLargeDataProcessing() {
    std::cout << "\n========== 大数据量处理演示 ==========" << std::endl;
    
    auto workbook = std::make_shared<Workbook>();
    SharedStringTable sst;
    FormatPool format_pool;
    
    auto worksheet = std::make_unique<Worksheet>("大数据测试", workbook, 1);
    worksheet->setSharedStringTable(&sst);
    worksheet->setFormatPool(&format_pool);
    worksheet->setOptimizeMode(true);
    
    const int LARGE_ROWS = 20000;
    const int LARGE_COLS = 15;
    
    std::cout << "处理 " << LARGE_ROWS << " x " << LARGE_COLS << " = " << (LARGE_ROWS * LARGE_COLS) << " 个单元格..." << std::endl;
    
    {
        PerformanceTimer timer("大数据量写入");
        
        for (int row = 0; row < LARGE_ROWS; ++row) {
            for (int col = 0; col < LARGE_COLS; ++col) {
                if (col % 3 == 0) {
                    worksheet->writeString(row, col, "大数据_" + std::to_string(row % 50));
                } else if (col % 3 == 1) {
                    worksheet->writeNumber(row, col, row * col * 0.001);
                } else {
                    worksheet->writeBoolean(row, col, (row + col) % 2 == 0);
                }
            }
            
            // 每5000行输出进度并刷新
            if (row % 5000 == 0 && row > 0) {
                worksheet->flushCurrentRow();
                std::cout << "已处理 " << row << " 行..." << std::endl;
            }
        }
        
        // 刷新最后一行
        worksheet->flushCurrentRow();
    }
    
    auto stats = worksheet->getPerformanceStats();
    std::cout << "\n大数据量处理结果:" << std::endl;
    std::cout << "  总单元格数: " << stats.total_cells << std::endl;
    std::cout << "  内存使用: " << formatMemory(stats.memory_usage) << std::endl;
    std::cout << "  平均每单元格: " << (stats.memory_usage / stats.total_cells) << " bytes" << std::endl;
    std::cout << "  SST字符串数: " << stats.sst_strings << std::endl;
    std::cout << "  SST压缩比: " << std::fixed << std::setprecision(1) << stats.sst_compression_ratio << "%" << std::endl;
    std::cout << "  唯一格式数: " << stats.unique_formats << std::endl;
    std::cout << "  格式去重比: " << std::fixed << std::setprecision(1) << stats.format_deduplication_ratio << "%" << std::endl;
}

// 演示动态模式切换
void demonstrateModeSwitching() {
    std::cout << "\n========== 动态模式切换演示 ==========" << std::endl;
    
    auto workbook = std::make_shared<Workbook>();
    SharedStringTable sst;
    FormatPool format_pool;
    
    auto worksheet = std::make_unique<Worksheet>("模式切换测试", workbook, 1);
    worksheet->setSharedStringTable(&sst);
    worksheet->setFormatPool(&format_pool);
    
    // 标准模式写入
    std::cout << "标准模式写入数据..." << std::endl;
    for (int i = 0; i < 100; ++i) {
        worksheet->writeString(i, 0, "标准_" + std::to_string(i));
        worksheet->writeNumber(i, 1, i * 1.5);
    }
    
    auto stats1 = worksheet->getPerformanceStats();
    std::cout << "标准模式 - 内存使用: " << formatMemory(stats1.memory_usage) << std::endl;
    
    // 切换到优化模式
    std::cout << "切换到优化模式..." << std::endl;
    worksheet->setOptimizeMode(true);
    
    // 优化模式写入
    for (int i = 100; i < 200; ++i) {
        worksheet->writeString(i, 0, "优化_" + std::to_string(i));
        worksheet->writeNumber(i, 1, i * 2.0);
    }
    worksheet->flushCurrentRow();
    
    auto stats2 = worksheet->getPerformanceStats();
    std::cout << "优化模式 - 内存使用: " << formatMemory(stats2.memory_usage) << std::endl;
    
    // 切换回标准模式
    std::cout << "切换回标准模式..." << std::endl;
    worksheet->setOptimizeMode(false);
    
    // 继续写入
    for (int i = 200; i < 300; ++i) {
        worksheet->writeString(i, 0, "标准2_" + std::to_string(i));
        worksheet->writeNumber(i, 1, i * 0.8);
    }
    
    auto stats3 = worksheet->getPerformanceStats();
    std::cout << "最终 - 总单元格: " << stats3.total_cells << ", 内存使用: " << formatMemory(stats3.memory_usage) << std::endl;
}

// 演示XML生成
void demonstrateXMLGeneration() {
    std::cout << "\n========== XML生成演示 ==========" << std::endl;
    
    auto workbook = std::make_shared<Workbook>();
    SharedStringTable sst;
    FormatPool format_pool;
    
    auto worksheet = std::make_unique<Worksheet>("XML测试", workbook, 1);
    worksheet->setSharedStringTable(&sst);
    worksheet->setFormatPool(&format_pool);
    worksheet->setOptimizeMode(true);
    
    // 创建格式
    auto header_format = std::make_shared<Format>();
    header_format->setBold(true);
    header_format->setBackgroundColor(0xCCCCCC);
    
    // 添加数据
    worksheet->writeString(0, 0, "姓名", header_format);
    worksheet->writeString(0, 1, "年龄", header_format);
    worksheet->writeString(0, 2, "城市", header_format);
    
    std::vector<std::string> names = {"张三", "李四", "王五"};
    std::vector<int> ages = {25, 30, 28};
    std::vector<std::string> cities = {"北京", "上海", "广州"};
    
    for (size_t i = 0; i < names.size(); ++i) {
        int row = i + 1;
        worksheet->writeString(row, 0, names[i]);
        worksheet->writeNumber(row, 1, ages[i]);
        worksheet->writeString(row, 2, cities[i]);
    }
    
    worksheet->flushCurrentRow();
    
    // 生成XML
    {
        PerformanceTimer timer("XML生成");
        std::string xml = worksheet->generateXML();
        std::cout << "生成的XML长度: " << xml.length() << " 字符" << std::endl;
        
        // 显示XML的前300个字符作为示例
        std::cout << "\nXML示例 (前300字符):" << std::endl;
        std::cout << xml.substr(0, 300) << "..." << std::endl;
    }
}

int main() {
    std::cout << "========================================" << std::endl;
    std::cout << "    Worksheet 集成优化功能演示" << std::endl;
    std::cout << "========================================" << std::endl;
    
    try {
        // 演示各个优化功能
        demonstrateBasicOptimization();
        demonstrateOptimizeMode();
        demonstrateLargeDataProcessing();
        demonstrateModeSwitching();
        demonstrateXMLGeneration();
        
        std::cout << "\n========================================" << std::endl;
        std::cout << "           演示完成!" << std::endl;
        std::cout << "========================================" << std::endl;
        
    } catch (const std::exception& e) {
        std::cerr << "错误: " << e.what() << std::endl;
        return 1;
    }
    
    return 0;
}