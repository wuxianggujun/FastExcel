#include <iostream>
#include <chrono>
#include "fastexcel/core/Workbook.hpp"

using namespace fastexcel::core;

void testMode(const std::string& filename, WorkbookMode mode, const std::string& mode_name) {
    std::cout << "\n=== Testing " << mode_name << " Mode ===" << std::endl;
    
    auto start = std::chrono::high_resolution_clock::now();
    
    // 创建工作簿
    auto workbook = Workbook::create(filename);
    workbook->open();
    
    // 设置模式
    workbook->setMode(mode);
    
    // 创建工作表并添加数据
    auto worksheet = workbook->addWorksheet("TestSheet");
    
    // 添加一些测试数据
    const int rows = 1000;
    const int cols = 10;
    
    for (int row = 0; row < rows; ++row) {
        for (int col = 0; col < cols; ++col) {
            if (col == 0) {
                worksheet->writeString(row, col, "Row " + std::to_string(row + 1));
            } else {
                worksheet->writeNumber(row, col, row * cols + col);
            }
        }
    }
    
    // 保存文件
    workbook->save();
    workbook->close();
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);
    
    std::cout << "File created: " << filename << std::endl;
    std::cout << "Time taken: " << duration.count() << " ms" << std::endl;
    std::cout << "Total cells: " << rows * cols << std::endl;
}

void testAutoModeWithDifferentThresholds() {
    std::cout << "\n=== Testing Auto Mode with Different Thresholds ===" << std::endl;
    
    auto workbook = Workbook::create("test_auto_mode_custom.xlsx");
    workbook->open();
    
    // 设置自定义阈值（非常低，强制使用流式模式）
    workbook->setAutoModeThresholds(100, 1024 * 1024); // 100个单元格或1MB
    workbook->setMode(WorkbookMode::AUTO);
    
    auto worksheet = workbook->addWorksheet("AutoTest");
    
    // 添加200个单元格（超过阈值）
    for (int i = 0; i < 200; ++i) {
        worksheet->writeNumber(i, 0, i);
    }
    
    std::cout << "Creating file with 200 cells (threshold: 100 cells)" << std::endl;
    std::cout << "Expected: Should use streaming mode" << std::endl;
    
    workbook->save();
    workbook->close();
}

int main() {
    try {
        std::cout << "FastExcel Mode Selection Test" << std::endl;
        std::cout << "=============================" << std::endl;
        
        // 测试不同模式
        testMode("test_auto_mode.xlsx", WorkbookMode::AUTO, "Auto");
        testMode("test_batch_mode.xlsx", WorkbookMode::BATCH, "Batch");
        testMode("test_streaming_mode.xlsx", WorkbookMode::STREAMING, "Streaming");
        
        // 测试自定义阈值
        testAutoModeWithDifferentThresholds();
        
        // 测试向后兼容性
        std::cout << "\n=== Testing Backward Compatibility ===" << std::endl;
        auto workbook = Workbook::create("test_backward_compat.xlsx");
        workbook->open();
        
        // 使用旧的API
        workbook->setStreamingXML(false); // 应该设置为BATCH模式
        std::cout << "Called setStreamingXML(false)" << std::endl;
        std::cout << "Current mode: " << (workbook->getMode() == WorkbookMode::BATCH ? "BATCH" : "OTHER") << std::endl;
        
        workbook->setStreamingXML(true); // 应该设置为STREAMING模式
        std::cout << "Called setStreamingXML(true)" << std::endl;
        std::cout << "Current mode: " << (workbook->getMode() == WorkbookMode::STREAMING ? "STREAMING" : "OTHER") << std::endl;
        
        auto worksheet = workbook->addWorksheet("CompatTest");
        worksheet->writeString(0, 0, "Backward compatibility test");
        
        workbook->save();
        workbook->close();
        
        std::cout << "\nAll tests completed successfully!" << std::endl;
        
    } catch (const std::exception& e) {
        std::cerr << "Error: " << e.what() << std::endl;
        return 1;
    }
    
    return 0;
}