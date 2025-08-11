#include "fastexcel/FastExcel.hpp"
#include <iostream>

using namespace fastexcel;
using namespace fastexcel::core;

int main() {
    try {
        // 创建一个工作簿
        auto workbook = Workbook::create(Path("test_format.xlsx"));
        workbook->open();
        
        auto worksheet = workbook->addSheet("Test");
        
        // 获取一个单元格并设置格式
        auto& cell = worksheet->getCell(0, 0);
        
        std::cout << "Before setting format:" << std::endl;
        std::cout << "  hasFormat(): " << (cell.hasFormat() ? "true" : "false") << std::endl;
        std::cout << "  isEmpty(): " << (cell.isEmpty() ? "true" : "false") << std::endl;
        
        // 创建一个简单的格式
        FormatDescriptor format;
        format.fill.pattern_type = FillPatternType::Solid;
        format.fill.fg_color = Color::fromRGB(255, 255, 0); // 黄色背景
        
        // 通过FormatRepository添加格式
        int style_id = workbook->addStyle(format);
        std::cout << "Added style with ID: " << style_id << std::endl;
        
        // 获取格式并设置到单元格
        auto format_desc = workbook->getStyle(style_id);
        if (format_desc) {
            cell.setFormat(format_desc);
            std::cout << "Format set successfully" << std::endl;
        }
        
        std::cout << "After setting format:" << std::endl;
        std::cout << "  hasFormat(): " << (cell.hasFormat() ? "true" : "false") << std::endl;
        std::cout << "  isEmpty(): " << (cell.isEmpty() ? "true" : "false") << std::endl;
        
        // 检查格式是否被正确获取
        auto retrieved_format = cell.getFormatDescriptor();
        if (retrieved_format) {
            std::cout << "  getFormatDescriptor(): valid" << std::endl;
        } else {
            std::cout << "  getFormatDescriptor(): null" << std::endl;
        }
        
        return 0;
        
    } catch (const std::exception& e) {
        std::cerr << "Error: " << e.what() << std::endl;
        return 1;
    }
}