#include "fastexcel/core/Workbook.hpp"
#include <iostream>

int main() {
    try {
        // 创建最简单的空工作簿（仅包含一个空工作表）
        auto workbook = fastexcel::core::Workbook::create("simple_test.xlsx");
        
        // 禁用共享字符串，使用内联字符串
        workbook->setUseSharedStrings(false);
        
        if (!workbook->open()) {
            std::cerr << "Failed to open workbook" << std::endl;
            return 1;
        }
        
        // 只添加一个空的工作表 Sheet1（不写入任何数据）
        auto sheet1 = workbook->addWorksheet("Sheet1");
        if (!sheet1) {
            std::cerr << "Failed to create Sheet1" << std::endl;
            return 1;
        }
        
        // 不写入任何数据，保持工作表完全为空
        
        // 保存文件
        if (!workbook->save()) {
            std::cerr << "Failed to save workbook" << std::endl;
            return 1;
        }
        
        workbook->close();
        
        std::cout << "Empty test file created successfully: simple_test.xlsx" << std::endl;
        std::cout << "- Contains only one empty Sheet1" << std::endl;
        std::cout << "- Shared strings disabled (using inline strings)" << std::endl;
        
        return 0;
        
    } catch (const std::exception& e) {
        std::cerr << "Exception: " << e.what() << std::endl;
        return 1;
    }
}