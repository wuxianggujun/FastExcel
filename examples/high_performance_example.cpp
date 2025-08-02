/**
 * @file high_performance_example.cpp
 * @brief 高性能Excel文件生成示例
 * 
 * 演示FastExcel库的性能优化功能：
 * - 流式XML写入
 * - 禁用SharedStrings
 * - 优化压缩级别
 * - 行缓冲机制
 */

#include "fastexcel/FastExcel.hpp"
#include <iostream>
#include <chrono>
#include <iomanip>

using namespace fastexcel::core;

void demonstrateHighPerformanceMode() {
    std::cout << "=== 高性能模式演示 ===" << std::endl;
    
    // 创建工作簿
    auto workbook = Workbook::create("high_performance_test.xlsx");
    if (!workbook->open()) {
        std::cerr << "无法打开工作簿" << std::endl;
        return;
    }
    
    // 启用高性能模式
    workbook->setHighPerformanceMode(true);
    
    // 添加工作表
    auto worksheet = workbook->addWorksheet("PerformanceTest");
    if (!worksheet) {
        std::cerr << "无法创建工作表" << std::endl;
        return;
    }
    
    std::cout << "开始生成大量数据..." << std::endl;
    auto start_time = std::chrono::high_resolution_clock::now();
    
    // 生成大量数据（100,000行 x 10列）
    const int rows = 100000;
    const int cols = 10;
    
    // 写入表头
    for (int col = 0; col < cols; ++col) {
        worksheet->writeString(0, col, "Column " + std::to_string(col + 1));
    }
    
    // 写入数据
    for (int row = 1; row <= rows; ++row) {
        for (int col = 0; col < cols; ++col) {
            if (col % 3 == 0) {
                // 数字数据
                worksheet->writeNumber(row, col, row * col * 1.5);
            } else if (col % 3 == 1) {
                // 字符串数据（在高性能模式下不使用SharedStrings）
                worksheet->writeString(row, col, "Data_" + std::to_string(row) + "_" + std::to_string(col));
            } else {
                // 布尔数据
                worksheet->writeBoolean(row, col, (row + col) % 2 == 0);
            }
        }
        
        // 每10000行显示进度
        if (row % 10000 == 0) {
            std::cout << "已处理 " << row << " 行..." << std::endl;
        }
    }
    
    auto data_time = std::chrono::high_resolution_clock::now();
    auto data_duration = std::chrono::duration_cast<std::chrono::milliseconds>(data_time - start_time);
    std::cout << "数据写入完成，耗时: " << data_duration.count() << " ms" << std::endl;
    
    // 保存文件
    std::cout << "开始保存文件..." << std::endl;
    auto save_start = std::chrono::high_resolution_clock::now();
    
    if (workbook->save()) {
        auto save_end = std::chrono::high_resolution_clock::now();
        auto save_duration = std::chrono::duration_cast<std::chrono::milliseconds>(save_end - save_start);
        auto total_duration = std::chrono::duration_cast<std::chrono::milliseconds>(save_end - start_time);
        
        std::cout << "文件保存成功！" << std::endl;
        std::cout << "保存耗时: " << save_duration.count() << " ms" << std::endl;
        std::cout << "总耗时: " << total_duration.count() << " ms" << std::endl;
        
        // 计算性能指标
        double total_cells = static_cast<double>(rows * cols);
        double cells_per_second = total_cells / (total_duration.count() / 1000.0);
        
        std::cout << std::fixed << std::setprecision(0);
        std::cout << "总单元格数: " << total_cells << std::endl;
        std::cout << "处理速度: " << cells_per_second << " 单元格/秒" << std::endl;
        
    } else {
        std::cerr << "文件保存失败！" << std::endl;
    }
    
    workbook->close();
}

void demonstrateCustomPerformanceSettings() {
    std::cout << "\n=== 自定义性能设置演示 ===" << std::endl;
    
    auto workbook = Workbook::create("custom_performance_test.xlsx");
    if (!workbook->open()) {
        std::cerr << "无法打开工作簿" << std::endl;
        return;
    }
    
    // 自定义性能设置
    workbook->setUseSharedStrings(false);        // 禁用共享字符串
    workbook->setStreamingXML(true);             // 启用流式XML
    workbook->setRowBufferSize(2000);            // 设置行缓冲大小
    workbook->setCompressionLevel(3);            // 中等压缩级别
    workbook->setXMLBufferSize(2 * 1024 * 1024); // 2MB XML缓冲区
    
    std::cout << "性能设置:" << std::endl;
    std::cout << "- SharedStrings: 禁用" << std::endl;
    std::cout << "- StreamingXML: 启用" << std::endl;
    std::cout << "- RowBufferSize: 2000" << std::endl;
    std::cout << "- CompressionLevel: 3" << std::endl;
    std::cout << "- XMLBufferSize: 2MB" << std::endl;
    
    auto worksheet = workbook->addWorksheet("CustomSettings");
    if (!worksheet) {
        std::cerr << "无法创建工作表" << std::endl;
        return;
    }
    
    auto start_time = std::chrono::high_resolution_clock::now();
    
    // 生成中等规模数据（50,000行 x 5列）
    const int rows = 50000;
    const int cols = 5;
    
    for (int row = 0; row < rows; ++row) {
        for (int col = 0; col < cols; ++col) {
            if (col == 0) {
                worksheet->writeString(row, col, "Item_" + std::to_string(row));
            } else if (col == 1) {
                worksheet->writeNumber(row, col, row * 10.5);
            } else if (col == 2) {
                worksheet->writeFormula(row, col, "B" + std::to_string(row + 1) + "*2");
            } else if (col == 3) {
                worksheet->writeBoolean(row, col, row % 2 == 0);
            } else {
                worksheet->writeString(row, col, "Status_" + std::to_string(row % 10));
            }
        }
        
        if (row % 5000 == 0 && row > 0) {
            std::cout << "已处理 " << row << " 行..." << std::endl;
        }
    }
    
    if (workbook->save()) {
        auto end_time = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time);
        
        double total_cells = static_cast<double>(rows * cols);
        double cells_per_second = total_cells / (duration.count() / 1000.0);
        
        std::cout << "自定义设置测试完成！" << std::endl;
        std::cout << "总耗时: " << duration.count() << " ms" << std::endl;
        std::cout << std::fixed << std::setprecision(0);
        std::cout << "处理速度: " << cells_per_second << " 单元格/秒" << std::endl;
    } else {
        std::cerr << "文件保存失败！" << std::endl;
    }
    
    workbook->close();
}

void demonstrateCompressionLevels() {
    std::cout << "\n=== 压缩级别对比演示 ===" << std::endl;
    
    const int rows = 10000;
    const int cols = 8;
    
    // 测试不同压缩级别
    std::vector<int> compression_levels = {0, 1, 3, 6, 9};
    
    for (int level : compression_levels) {
        std::string filename = "compression_test_level_" + std::to_string(level) + ".xlsx";
        
        auto workbook = Workbook::create(filename);
        if (!workbook->open()) {
            std::cerr << "无法打开工作簿: " << filename << std::endl;
            continue;
        }
        
        // 设置压缩级别
        workbook->setCompressionLevel(level);
        workbook->setUseSharedStrings(false); // 禁用共享字符串以获得一致的测试结果
        
        auto worksheet = workbook->addWorksheet("CompressionTest");
        
        auto start_time = std::chrono::high_resolution_clock::now();
        
        // 生成相同的数据
        for (int row = 0; row < rows; ++row) {
            for (int col = 0; col < cols; ++col) {
                if (col % 2 == 0) {
                    worksheet->writeString(row, col, "TestData_" + std::to_string(row) + "_" + std::to_string(col));
                } else {
                    worksheet->writeNumber(row, col, row * col * 0.123456789);
                }
            }
        }
        
        if (workbook->save()) {
            auto end_time = std::chrono::high_resolution_clock::now();
            auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time);
            
            std::cout << "压缩级别 " << level << ": " << duration.count() << " ms" << std::endl;
        } else {
            std::cerr << "保存失败: " << filename << std::endl;
        }
        
        workbook->close();
    }
}

int main() {
    std::cout << "FastExcel 高性能示例程序" << std::endl;
    std::cout << "=========================" << std::endl;
    
    try {
        // 演示高性能模式
        demonstrateHighPerformanceMode();
        
        // 演示自定义性能设置
        demonstrateCustomPerformanceSettings();
        
        // 演示压缩级别对比
        demonstrateCompressionLevels();
        
        std::cout << "\n所有测试完成！" << std::endl;
        std::cout << "生成的文件:" << std::endl;
        std::cout << "- high_performance_test.xlsx (高性能模式)" << std::endl;
        std::cout << "- custom_performance_test.xlsx (自定义设置)" << std::endl;
        std::cout << "- compression_test_level_*.xlsx (压缩级别测试)" << std::endl;
        
    } catch (const std::exception& e) {
        std::cerr << "发生异常: " << e.what() << std::endl;
        return 1;
    }
    
    return 0;
}