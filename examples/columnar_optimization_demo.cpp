/**
 * @file columnar_optimization_demo.cpp
 * @brief 列式存储优化演示程序 - 对比内存使用情况
 * 
 * 本程序演示 FastExcel 列式存储模式相比传统 Cell 对象模式的内存优化效果。
 * 
 * 核心优化原理：
 * 1. 传统模式：每个单元格创建一个 Cell 对象，包含值、格式、公式等完整信息
 * 2. 列式模式：数据按列分类存储，完全跳过 Cell 对象创建，直接使用 SST 索引
 * 
 * 预期优化效果：
 * - 内存使用减少 60-80%
 * - 解析速度提升 3-5倍
 * - 适合只读场景的大文件处理
 */

#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/WorkbookTypes.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <iostream>
#include <chrono>
#include <iomanip>
#include <variant>
#include <algorithm>

using namespace fastexcel;
using namespace std::chrono;

/**
 * @brief 格式化内存大小显示
 */
std::string formatMemorySize(size_t bytes) {
    if (bytes >= 1024 * 1024 * 1024) {
        return std::to_string(bytes / (1024 * 1024 * 1024)) + " GB";
    } else if (bytes >= 1024 * 1024) {
        return std::to_string(bytes / (1024 * 1024)) + " MB";  
    } else if (bytes >= 1024) {
        return std::to_string(bytes / 1024) + " KB";
    } else {
        return std::to_string(bytes) + " B";
    }
}

/**
 * @brief 测试传统 Cell 对象模式
 */
void testTraditionalMode(const std::string& filepath) {
    std::cout << "\n=== 传统 Cell 对象模式测试 ===" << std::endl;
    
    auto start_time = high_resolution_clock::now();
    
    // 使用标准模式打开文件
    auto workbook = core::Workbook::openReadOnly(filepath);
    if (!workbook) {
        std::cout << "❌ 无法打开文件: " << filepath << std::endl;
        return;
    }
    
    auto end_time = high_resolution_clock::now();
    auto duration = duration_cast<milliseconds>(end_time - start_time);
    
    auto worksheet = workbook->getSheet(0);
    if (!worksheet) {
        std::cout << "❌ 无法获取工作表" << std::endl;
        return;
    }
    
    // 获取使用范围
    auto used_range = worksheet->getUsedRange();
    int rows = used_range.first;
    int cols = used_range.second;
    auto used_range_full = worksheet->getUsedRangeFull();
    int first_row = std::get<0>(used_range_full);
    int first_col = std::get<1>(used_range_full);
    int last_row = std::get<2>(used_range_full);
    int last_col = std::get<3>(used_range_full);
    
    // 计算总单元格数
    int total_cells = 0;
    for (int row = first_row; row <= last_row; ++row) {
        for (int col = first_col; col <= last_col; ++col) {
            if (worksheet->hasCellAt(row, col)) {
                total_cells++;
            }
        }
    }
    
    // 获取内存统计
    auto perf_stats = worksheet->getPerformanceStats();
    
    std::cout << "📊 解析耗时: " << duration.count() << " ms" << std::endl;
    std::cout << "📊 工作表范围: " << (last_row + 1) << " 行 × " << (last_col + 1) << " 列" << std::endl;
    std::cout << "📊 有效单元格: " << total_cells << " 个" << std::endl;
    std::cout << "📊 内存使用: " << formatMemorySize(perf_stats.memory_usage) << std::endl;
    std::cout << "📊 共享字符串: " << perf_stats.sst_strings << " 个" << std::endl;
    std::cout << "📊 格式数量: " << perf_stats.unique_formats << " 个" << std::endl;
    
    // 随机采样几个单元格显示内容
    std::cout << "\n📋 数据采样 (前5行×5列):" << std::endl;
    int max_sample_row = std::min(first_row + 5, last_row + 1);
    int max_sample_col = std::min(first_col + 5, last_col + 1);
    for (int row = first_row; row < max_sample_row; ++row) {
        for (int col = first_col; col < max_sample_col; ++col) {
            if (worksheet->hasCellAt(row, col)) {
                auto& cell = worksheet->getCell(row, col);
                std::cout << "[" << row << "," << col << "]=" << cell.asString() << " ";
            }
        }
        std::cout << std::endl;
    }
}

/**
 * @brief 测试列式存储优化模式  
 */
void testColumnarMode(const std::string& filepath) {
    std::cout << "\n=== 列式存储优化模式测试 ===" << std::endl;
    
    // 配置列式存储选项
    core::WorkbookOptions options;
    options.enable_columnar_storage = true;
    // 可选：只读取指定列（列投影优化）
    // options.projected_columns = {0, 1, 2, 3, 4};  // 只读取前5列
    // 可选：限制读取行数
    // options.max_rows = 10000;  // 只读取前1万行
    
    auto start_time = high_resolution_clock::now();
    
    // 使用列式存储模式打开文件
    auto workbook = core::Workbook::openReadOnly(filepath, options);
    if (!workbook) {
        std::cout << "❌ 无法打开文件: " << filepath << std::endl;
        return;
    }
    
    auto end_time = high_resolution_clock::now();
    auto duration = duration_cast<milliseconds>(end_time - start_time);
    
    auto worksheet = workbook->getSheet(0);
    if (!worksheet) {
        std::cout << "❌ 无法获取工作表" << std::endl;
        return;
    }
    
    std::cout << "📊 解析耗时: " << duration.count() << " ms" << std::endl;
    std::cout << "📊 列式模式: " << (worksheet->isColumnarMode() ? "✅ 启用" : "❌ 未启用") << std::endl;
    
    if (worksheet->isColumnarMode()) {
        // 获取列式存储统计
        auto data_count = worksheet->getColumnarDataCount();
        auto memory_usage = worksheet->getColumnarMemoryUsage();
        
        std::cout << "📊 列式数据点: " << data_count << " 个" << std::endl;
        std::cout << "📊 列式内存: " << formatMemorySize(memory_usage) << std::endl;
        
        // 演示列式数据访问
        std::cout << "\n📋 列式数据采样:" << std::endl;
        
        // 遍历前5列的数据
        for (uint32_t col = 0; col < 5; ++col) {
            std::cout << "列 " << col << ": ";
            
            // 获取该列的数字数据
            auto number_data = worksheet->getNumberColumn(col);
            if (!number_data.empty()) {
                std::cout << "数字(" << number_data.size() << "个) ";
                
                // 显示前3个数值
                int count = 0;
                for (auto it = number_data.begin(); it != number_data.end() && count < 3; ++it, ++count) {
                    std::cout << "[" << it->first << "]=" << it->second << " ";
                }
            }
            
            // 获取该列的字符串数据（SST索引）
            auto string_data = worksheet->getStringColumn(col);
            if (!string_data.empty()) {
                std::cout << "字符串(" << string_data.size() << "个) ";
                
                // 显示前3个SST索引
                int count = 0;
                for (auto it = string_data.begin(); it != string_data.end() && count < 3; ++it, ++count) {
                    std::cout << "[" << it->first << "]=SST#" << it->second << " ";
                }
            }
            
            // 获取该列的布尔数据
            auto boolean_data = worksheet->getBooleanColumn(col);
            if (!boolean_data.empty()) {
                std::cout << "布尔(" << boolean_data.size() << "个) ";
            }
            
            // 获取该列的错误数据（内联字符串等）
            auto error_data = worksheet->getErrorColumn(col);
            if (!error_data.empty()) {
                std::cout << "文本(" << error_data.size() << "个) ";
                
                // 显示前2个文本值
                int count = 0;
                for (auto it = error_data.begin(); it != error_data.end() && count < 2; ++it, ++count) {
                    std::string display_text = it->second.length() > 20 ? it->second.substr(0, 20) + "..." : it->second;
                    std::cout << "[" << it->first << "]=" << display_text << " ";
                }
            }
            
            std::cout << std::endl;
        }
        
        // 演示列遍历功能
        std::cout << "\n📋 列遍历演示 (第0列前5行):" << std::endl;
        int callback_count = 0;
        worksheet->forEachInColumn(0, [&callback_count](uint32_t row, const auto& value) {
            if (callback_count < 5) {
                std::cout << "行 " << row << ": 数据类型变体" << std::endl;
                callback_count++;
            }
        });
    }
}

/**
 * @brief 对比两种模式的性能差异
 */
void comparePerformance(const std::string& filepath) {
    std::cout << "\n=== 性能对比分析 ===" << std::endl;
    
    // 传统模式测试
    std::cout << "\n🔄 正在测试传统模式..." << std::endl;
    auto start1 = high_resolution_clock::now();
    auto workbook1 = core::Workbook::openReadOnly(filepath);
    auto end1 = high_resolution_clock::now();
    auto duration1 = duration_cast<milliseconds>(end1 - start1);
    
    size_t traditional_memory = 0;
    if (workbook1 && workbook1->getSheet(0)) {
        traditional_memory = workbook1->getSheet(0)->getPerformanceStats().memory_usage;
    }
    
    // 列式模式测试
    std::cout << "🔄 正在测试列式模式..." << std::endl;
    core::WorkbookOptions options;
    options.enable_columnar_storage = true;
    
    auto start2 = high_resolution_clock::now();
    auto workbook2 = core::Workbook::openReadOnly(filepath, options);
    auto end2 = high_resolution_clock::now();
    auto duration2 = duration_cast<milliseconds>(end2 - start2);
    
    size_t columnar_memory = 0;
    if (workbook2 && workbook2->getSheet(0) && workbook2->getSheet(0)->isColumnarMode()) {
        columnar_memory = workbook2->getSheet(0)->getColumnarMemoryUsage();
    }
    
    // 对比结果
    std::cout << "\n📈 性能对比结果:" << std::endl;
    std::cout << std::setw(20) << "指标" << std::setw(15) << "传统模式" << std::setw(15) << "列式模式" << std::setw(15) << "优化幅度" << std::endl;
    std::cout << std::string(65, '-') << std::endl;
    
    std::cout << std::setw(20) << "解析耗时" 
              << std::setw(15) << (std::to_string(duration1.count()) + " ms")
              << std::setw(15) << (std::to_string(duration2.count()) + " ms");
    if (duration1.count() > 0) {
        double speed_improvement = (double)duration1.count() / duration2.count();
        std::cout << std::setw(15) << (std::to_string(speed_improvement) + "x 加速");
    }
    std::cout << std::endl;
    
    std::cout << std::setw(20) << "内存使用" 
              << std::setw(15) << formatMemorySize(traditional_memory)
              << std::setw(15) << formatMemorySize(columnar_memory);
    if (traditional_memory > 0 && columnar_memory > 0) {
        double memory_reduction = (1.0 - (double)columnar_memory / traditional_memory) * 100;
        std::cout << std::setw(15) << (std::to_string((int)memory_reduction) + "% 减少");
    }
    std::cout << std::endl;
    
    std::cout << "\n💡 优化建议:" << std::endl;
    if (columnar_memory < traditional_memory) {
        std::cout << "✅ 列式存储有效减少了内存使用，适合大文件只读场景" << std::endl;
    }
    if (duration2.count() < duration1.count()) {
        std::cout << "✅ 列式存储提升了解析速度，适合快速数据加载" << std::endl;
    }
    std::cout << "✅ 建议在只读场景下使用列式存储模式以获得最佳性能" << std::endl;
    std::cout << "✅ 可配置列投影和行限制进一步优化内存和速度" << std::endl;
}

int main(int argc, char* argv[]) {
    // 获取文件路径
    std::string filepath = "C:\\Users\\wuxianggujun\\CodeSpace\\CMakeProjects\\FastExcel\\test_xlsx\\合并去年和今年的数据.xlsx";
    
    std::cout << "FastExcel 列式存储优化演示程序" << std::endl;
    std::cout << "===============================" << std::endl;
    std::cout << "测试文件: " << filepath << std::endl;
    
    try {
        // 测试传统模式
        testTraditionalMode(filepath);
        
        // 测试列式模式
        testColumnarMode(filepath);
        
        // 性能对比
        comparePerformance(filepath);
        
        std::cout << "\n🎉 演示完成！" << std::endl;
        
    } catch (const std::exception& e) {
        std::cout << "❌ 程序执行错误: " << e.what() << std::endl;
        return 1;
    }
    
    return 0;
}