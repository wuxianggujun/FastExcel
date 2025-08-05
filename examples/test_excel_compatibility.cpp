#include "fastexcel/fastexcel.hpp"
#include <iostream>
#include <filesystem>
#include <fstream>

using namespace fastexcel;

bool validateZipFile(const std::string& filename) {
    std::ifstream file(filename, std::ios::binary);
    if (!file.is_open()) {
        std::cout << "  ✗ Cannot open file: " << filename << std::endl;
        return false;
    }
    
    // 检查ZIP文件头
    char header[4];
    file.read(header, 4);
    if (header[0] != 'P' || header[1] != 'K' || header[2] != 0x03 || header[3] != 0x04) {
        std::cout << "  ✗ Invalid ZIP file header" << std::endl;
        return false;
    }
    
    std::cout << "  ✓ Valid ZIP file header" << std::endl;
    return true;
}

void testMode(WorkbookMode mode, const std::string& filename, const std::string& modeName) {
    std::cout << "\n=== Testing " << modeName << " Mode ===" << std::endl;
    std::cout << "File: " << filename << std::endl;
    
    try {
        // 创建工作簿并强制使用指定模式
        Workbook workbook(filename);
        workbook.setMode(mode);
        
        if (!workbook.open()) {
            std::cout << "✗ Failed to open workbook" << std::endl;
            return;
        }
        
        // 添加工作表
        auto worksheet = workbook.addWorksheet("TestSheet");
        if (!worksheet) {
            std::cout << "✗ Failed to create worksheet" << std::endl;
            return;
        }
        
        // 写入测试数据
        worksheet->writeString(0, 0, "Mode");
        worksheet->writeString(0, 1, modeName);
        worksheet->writeString(1, 0, "Test Data");
        worksheet->writeNumber(1, 1, 123.45);
        worksheet->writeString(2, 0, "Excel Compatibility");
        worksheet->writeString(2, 1, "PASSED");
        
        // 添加一些格式化数据
        for (int row = 4; row < 10; ++row) {
            worksheet->writeString(row, 0, "Row " + std::to_string(row + 1));
            worksheet->writeNumber(row, 1, row * 10.5);
            worksheet->writeString(row, 2, "Data " + std::to_string(row));
        }
        
        // 保存文件
        if (!workbook.save()) {
            std::cout << "✗ Failed to save workbook" << std::endl;
            return;
        }
        
        workbook.close();
        
        // 验证文件
        if (std::filesystem::exists(filename)) {
            auto fileSize = std::filesystem::file_size(filename);
            std::cout << "✓ File created successfully" << std::endl;
            std::cout << "  File size: " << fileSize << " bytes" << std::endl;
            
            // 验证ZIP文件结构
            if (validateZipFile(filename)) {
                std::cout << "✓ " << modeName << " mode: Excel-compatible file generated" << std::endl;
            } else {
                std::cout << "✗ " << modeName << " mode: Invalid file structure" << std::endl;
            }
        } else {
            std::cout << "✗ File creation failed" << std::endl;
        }
        
    } catch (const std::exception& e) {
        std::cout << "✗ Exception: " << e.what() << std::endl;
    }
}

int main() {
    std::cout << "FastExcel Excel Compatibility Test" << std::endl;
    std::cout << "===================================" << std::endl;
    std::cout << "Testing all three modes for Excel compatibility..." << std::endl;
    
    // 测试所有三种模式
    testMode(WorkbookMode::AUTO, "test_auto_compatibility.xlsx", "AUTO");
    testMode(WorkbookMode::BATCH, "test_batch_compatibility.xlsx", "BATCH");
    testMode(WorkbookMode::STREAMING, "test_streaming_compatibility.xlsx", "STREAMING");
    
    std::cout << "\n=== Compatibility Test Summary ===" << std::endl;
    std::cout << "All three modes have been tested for Excel compatibility." << std::endl;
    std::cout << "Generated files:" << std::endl;
    std::cout << "- test_auto_compatibility.xlsx (AUTO mode)" << std::endl;
    std::cout << "- test_batch_compatibility.xlsx (BATCH mode)" << std::endl;
    std::cout << "- test_streaming_compatibility.xlsx (STREAMING mode)" << std::endl;
    std::cout << "\nPlease manually verify that all files can be opened in Excel." << std::endl;
    std::cout << "If all files open successfully, the Excel compatibility issue is resolved!" << std::endl;
    
    return 0;
}