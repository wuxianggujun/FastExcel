#include "fastexcel/fastexcel.hpp"
#include <iostream>
#include <filesystem>

using namespace fastexcel;

void testMode(WorkbookMode mode, const std::string& filename, const std::string& modeName) {
    std::cout << "\n=== Testing " << modeName << " Mode ===" << std::endl;
    std::cout << "Creating workbook: " << filename << std::endl;
    
    try {
        // 创建工作簿并强制使用指定模式
        Workbook workbook(filename);
        workbook.setMode(mode);
        
        // 添加工作表
        auto worksheet = workbook.createWorksheet("TestSheet");
        
        // 写入测试数据
        worksheet->writeString(0, 0, "Mode");
        worksheet->writeString(0, 1, modeName);
        worksheet->writeString(1, 0, "Cell A2");
        worksheet->writeNumber(1, 1, 123.45);
        worksheet->writeString(2, 0, "Cell A3");
        worksheet->writeNumber(2, 1, 678.90);
        
        // 添加一些格式化数据
        for (int row = 4; row < 10; ++row) {
            worksheet->writeString(row, 0, "Row " + std::to_string(row + 1));
            worksheet->writeNumber(row, 1, row * 10.5);
            worksheet->writeString(row, 2, "Data " + std::to_string(row));
        }
        
        // 保存文件
        workbook.save();
        workbook.close();
        
        // 检查文件是否存在
        if (std::filesystem::exists(filename)) {
            auto fileSize = std::filesystem::file_size(filename);
            std::cout << "✓ File created successfully: " << filename << std::endl;
            std::cout << "  File size: " << fileSize << " bytes" << std::endl;
            
            // 基本的ZIP文件验证（检查文件头）
            std::ifstream file(filename, std::ios::binary);
            if (file.is_open()) {
                char header[4];
                file.read(header, 4);
                if (header[0] == 'P' && header[1] == 'K' && header[2] == 0x03 && header[3] == 0x04) {
                    std::cout << "  ✓ Valid ZIP file header detected" << std::endl;
                } else {
                    std::cout << "  ✗ Invalid ZIP file header" << std::endl;
                }
                file.close();
            }
        } else {
            std::cout << "✗ File creation failed: " << filename << std::endl;
        }
        
    } catch (const std::exception& e) {
        std::cout << "✗ Exception occurred: " << e.what() << std::endl;
    }
}

int main() {
    std::cout << "FastExcel Mode Compatibility Test" << std::endl;
    std::cout << "=================================" << std::endl;
    
    // 测试所有三种模式
    testMode(WorkbookMode::AUTO, "test_auto_mode.xlsx", "AUTO");
    testMode(WorkbookMode::BATCH, "test_batch_mode.xlsx", "BATCH");
    testMode(WorkbookMode::STREAMING, "test_streaming_mode.xlsx", "STREAMING");
    
    std::cout << "\n=== Test Summary ===" << std::endl;
    std::cout << "All three modes have been tested." << std::endl;
    std::cout << "Please manually verify that all generated .xlsx files can be opened in Excel." << std::endl;
    std::cout << "\nGenerated files:" << std::endl;
    std::cout << "- test_auto_mode.xlsx" << std::endl;
    std::cout << "- test_batch_mode.xlsx" << std::endl;
    std::cout << "- test_streaming_mode.xlsx" << std::endl;
    
    return 0;
}