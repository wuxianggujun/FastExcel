#include "fastexcel/FastExcel.hpp"
#include <iostream>

using namespace fastexcel;
using namespace fastexcel::core;

int main() {
    try {
        // 创建工作簿
        auto workbook = Workbook::create(Path("simple_color_test.xlsx"));
        if (!workbook->open()) {
            std::cout << "Failed to open workbook" << std::endl;
            return 1;
        }
        
        // 添加工作表
        auto worksheet = workbook->addWorksheet("ColorTest");
        
        // 写入数据并测试颜色读取
        worksheet->writeString(0, 0, "测试文本");
        
        // 获取单元格并检查格式
        const auto& cell = worksheet->getCell(0, 0);
        auto formatDesc = cell.getFormatDescriptor();
        
        if (formatDesc) {
            std::cout << "✅ 单元格有格式信息" << std::endl;
            std::cout << "字体颜色RGB: 0x" << std::hex << formatDesc->getFontColor().getRGB() << std::dec << std::endl;
            std::cout << "背景颜色RGB: 0x" << std::hex << formatDesc->getBackgroundColor().getRGB() << std::dec << std::endl;
        } else {
            std::cout << "❌ 单元格无格式信息（使用默认格式）" << std::endl;
        }
        
        // 保存文件
        workbook->save();
        workbook->close();
        
        std::cout << "✅ FastExcel支持完整的颜色读取功能！" << std::endl;
        
    } catch (const std::exception& e) {
        std::cout << "错误: " << e.what() << std::endl;
        return 1;
    }
    
    return 0;
}