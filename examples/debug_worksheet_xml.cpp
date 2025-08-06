#include <iostream>
#include <string>
#include <memory>
#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/Workbook.hpp"

int main() {
    try {
        // 创建工作簿和工作表
        auto workbook = fastexcel::core::Workbook::create("debug.xlsx");
        workbook->open();
        auto worksheet = workbook->addWorksheet("TestSheet");
        
        // 写入数据 - 和测试中完全相同
        worksheet->writeString(0, 0, "Hello");
        worksheet->writeNumber(0, 1, 123.45);
        
        // 检查单元格数据是否正确存储
        const auto& cell_00 = worksheet->getCell(0, 0);
        const auto& cell_01 = worksheet->getCell(0, 1);
        
        std::cout << "=== 单元格数据检查 ===" << std::endl;
        std::cout << "Cell (0,0) - isEmpty: " << cell_00.isEmpty() << std::endl;
        std::cout << "Cell (0,0) - isString: " << cell_00.isString() << std::endl;
        std::cout << "Cell (0,0) - value: '" << cell_00.getStringValue() << "'" << std::endl;
        std::cout << "Cell (0,0) - hasFormat: " << cell_00.hasFormat() << std::endl;
        
        std::cout << "Cell (0,1) - isEmpty: " << cell_01.isEmpty() << std::endl;
        std::cout << "Cell (0,1) - isNumber: " << cell_01.isNumber() << std::endl;
        std::cout << "Cell (0,1) - value: " << cell_01.getNumberValue() << std::endl;
        std::cout << "Cell (0,1) - hasFormat: " << cell_01.hasFormat() << std::endl;
        
        // 检查使用范围
        auto [max_row, max_col] = worksheet->getUsedRange();
        std::cout << "\n=== 使用范围 ===" << std::endl;
        std::cout << "Used Range: (" << max_row << ", " << max_col << ")" << std::endl;
        std::cout << "Cell count: " << worksheet->getCellCount() << std::endl;
        
        // 检查工作簿设置
        std::cout << "\n=== 工作簿设置 ===" << std::endl;
        auto options = workbook->getOptions();
        std::cout << "Use shared strings: " << options.use_shared_strings << std::endl;
        
        // 生成XML
        std::cout << "\n=== XML生成 ===" << std::endl;
        std::string xml;
        worksheet->generateXML([&xml](const char* data, size_t size) {
            xml.append(data, size);
        });
        
        std::cout << "XML length: " << xml.length() << std::endl;
        std::cout << "XML contains 'Hello': " << (xml.find("Hello") != std::string::npos) << std::endl;
        std::cout << "XML contains '<sheetData': " << (xml.find("<sheetData") != std::string::npos) << std::endl;
        std::cout << "XML contains '123.45': " << (xml.find("123.45") != std::string::npos) << std::endl;
        
        // 输出XML的前1000个字符
        std::cout << "\n=== XML内容预览 ===" << std::endl;
        std::cout << xml.substr(0, std::min(xml.length(), size_t(1000))) << std::endl;
        if (xml.length() > 1000) {
            std::cout << "... (truncated)" << std::endl;
        }
        
        return 0;
    } catch (const std::exception& e) {
        std::cerr << "Error: " << e.what() << std::endl;
        return 1;
    }
}