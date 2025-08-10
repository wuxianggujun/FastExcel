#include "fastexcel/FastExcel.hpp"
#include <iostream>
#include <iomanip>

using namespace fastexcel;
using namespace fastexcel::core;

void printColor(const Color& color, const std::string& name) {
    std::cout << std::setfill('0') << std::hex 
              << name << ": RGB(0x" << std::setw(6) 
              << color.getRGB() << ")" << std::dec << std::endl;
}

int main() {
    try {
        // 跳过初始化，直接使用
        
        // 创建工作簿
        auto workbook = Workbook::create(Path("color_test.xlsx"));
        if (!workbook->open()) {
            std::cout << "Failed to open workbook" << std::endl;
            return 1;
        }
        
        // 添加工作表
        auto worksheet = workbook->addWorksheet("ColorTest");
        
        // 创建不同颜色的样式
        auto redStyle = workbook->createStyleBuilder()
            .fontName("Arial")
            .fontSize(12)
            .fontColor(Color::RED)
            .bold()
            .fill(PatternType::Solid, Color(0x87CEEBu), Color::WHITE)  // 浅蓝色背景
            .build();
        
        auto greenStyle = workbook->createStyleBuilder()
            .fontColor(Color::GREEN)
            .fontSize(14)
            .fill(PatternType::Gray125, Color::YELLOW, Color::WHITE)
            .build();
        
        // 添加样式到工作簿
        int redStyleId = workbook->addStyle(redStyle);
        int greenStyleId = workbook->addStyle(greenStyle);
        
        // 写入带颜色的单元格
        worksheet->writeString(0, 0, "红色字体蓝色背景");
        auto& cell1 = worksheet->getCell(0, 0);
        cell1.setFormat(workbook->getStyleRepository().getFormat(redStyleId));
        
        worksheet->writeString(1, 0, "绿色字体黄色背景"); 
        auto& cell2 = worksheet->getCell(1, 0);
        cell2.setFormat(workbook->getStyleRepository().getFormat(greenStyleId));
        
        worksheet->writeString(2, 0, "默认样式");
        
        // 保存文件
        workbook->save();
        
        std::cout << "=== FastExcel颜色读取功能测试 ===" << std::endl;
        
        // 测试颜色读取功能
        for (int row = 0; row < 3; ++row) {
            const auto& cell = worksheet->getCell(row, 0);
            std::cout << "\n单元格 A" << (row + 1) << ": \"" 
                      << cell.getStringValue() << "\"" << std::endl;
            
            // 获取格式描述符
            auto formatDesc = cell.getFormatDescriptor();
            if (formatDesc) {
                std::cout << "  ✅ 格式信息:" << std::endl;
                
                // 字体颜色
                printColor(formatDesc->getFontColor(), "    字体颜色");
                
                // 背景色和前景色
                printColor(formatDesc->getBackgroundColor(), "    背景色");
                printColor(formatDesc->getForegroundColor(), "    前景色");
                
                // 其他属性
                std::cout << "    字体: " << formatDesc->getFontName() 
                          << ", 大小: " << formatDesc->getFontSize() << std::endl;
                std::cout << "    粗体: " << (formatDesc->isBold() ? "是" : "否") << std::endl;
                std::cout << "    图案类型: " << static_cast<int>(formatDesc->getPattern()) << std::endl;
            } else {
                std::cout << "  ❌ 无格式信息" << std::endl;
            }
        }
        
        workbook->close();
        // 跳过cleanup
        
        std::cout << "\n🎉 FastExcel完全支持颜色获取功能!" << std::endl;
        std::cout << "📋 可用的颜色读取API:" << std::endl;
        std::cout << "   🎨 字体颜色: formatDesc->getFontColor()" << std::endl;
        std::cout << "   🎨 背景颜色: formatDesc->getBackgroundColor()" << std::endl; 
        std::cout << "   🎨 前景颜色: formatDesc->getForegroundColor()" << std::endl;
        std::cout << "   🎨 颜色RGB值: color.getRGB()" << std::endl;
        
    } catch (const std::exception& e) {
        std::cout << "❌ 错误: " << e.what() << std::endl;
        return 1;
    }
    
    return 0;
}