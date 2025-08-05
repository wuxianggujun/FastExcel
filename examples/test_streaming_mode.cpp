#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/FastExcel.hpp"
#include <iostream>

int main() {
    // 初始化FastExcel库并启用详细日志
    if (!fastexcel::initialize("logs/test_streaming_mode.log", true)) {
        std::cerr << "Failed to initialize FastExcel library" << std::endl;
        return 1;
    }
    
    try {
        std::cout << "=== Testing Streaming Mode ===" << std::endl;
        std::cout << "Creating workbook with forced streaming mode..." << std::endl;
        
        // 创建工作簿
        auto workbook = fastexcel::core::Workbook::create("test_streaming_mode.xlsx");
        
        // 强制使用流式模式
        workbook->setStreamingXML(true);
        workbook->setUseSharedStrings(false);  // 禁用共享字符串以简化测试
        
        if (!workbook->open()) {
            std::cerr << "Failed to open workbook" << std::endl;
            return 1;
        }
        
        // 添加一个工作表并写入一些数据
        auto sheet1 = workbook->addWorksheet("Sheet1");
        if (!sheet1) {
            std::cerr << "Failed to create Sheet1" << std::endl;
            return 1;
        }
        
        // 写入一些测试数据以触发流式模式
        std::cout << "Writing test data..." << std::endl;
        for (int row = 0; row < 10; ++row) {
            for (int col = 0; col < 5; ++col) {
                sheet1->writeNumber(row, col, row * 10 + col);
            }
        }
        
        // 保存文件
        std::cout << "Saving file with streaming mode..." << std::endl;
        if (!workbook->save()) {
            std::cerr << "Failed to save workbook" << std::endl;
            return 1;
        }
        
        workbook->close();
        
        std::cout << "Test file created successfully: test_streaming_mode.xlsx" << std::endl;
        std::cout << "- Forced streaming mode enabled" << std::endl;
        std::cout << "- Contains 10 rows x 5 columns of numeric data" << std::endl;
        std::cout << "- Shared strings disabled" << std::endl;
        std::cout << std::endl;
        std::cout << "=== Debug Information ===" << std::endl;
        std::cout << "Check logs/test_streaming_mode.log for detailed ZIP creation debug info" << std::endl;
        std::cout << "Look for 'STREAMING' sections to see how files are added to ZIP" << std::endl;
        std::cout << "All entries should now use flag: 0x0000 (no DATA_DESCRIPTOR)" << std::endl;
        
        return 0;
        
    } catch (const std::exception& e) {
        std::cerr << "Exception: " << e.what() << std::endl;
        return 1;
    }
}