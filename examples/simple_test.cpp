#include "fastexcel/core/Workbook.hpp"
#include <iostream>

int main() {
    try {
        // 创建最简单的工作簿
        auto workbook = fastexcel::core::Workbook::create("simple_test.xlsx");
        
        if (!workbook->open()) {
            std::cerr << "Failed to open workbook" << std::endl;
            return 1;
        }
        
        // 添加第一个工作表 Sheet1
        auto sheet1 = workbook->addWorksheet("Sheet1");
        if (!sheet1) {
            std::cerr << "Failed to create Sheet1" << std::endl;
            return 1;
        }
        
        // 在A1单元格写入"Hello"
        sheet1->writeString(0, 0, "Hello");
        
        // 添加第二个工作表 Sheet2
        auto sheet2 = workbook->addWorksheet("Sheet2");
        if (!sheet2) {
            std::cerr << "Failed to create Sheet2" << std::endl;
            return 1;
        }
        
        // 保存文件
        if (!workbook->save()) {
            std::cerr << "Failed to save workbook" << std::endl;
            return 1;
        }
        
        workbook->close();
        
        std::cout << "Simple test file created successfully: simple_test.xlsx" << std::endl;
        std::cout << "- Sheet1 with 'Hello' in A1" << std::endl;
        std::cout << "- Sheet2 (empty)" << std::endl;
        
        return 0;
        
    } catch (const std::exception& e) {
        std::cerr << "Exception: " << e.what() << std::endl;
        return 1;
    }
}