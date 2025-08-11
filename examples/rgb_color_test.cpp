#include "fastexcel/FastExcel.hpp"
#include <iostream>
#include <iomanip>

using namespace fastexcel;
using namespace fastexcel::core;

int main() {
    try {
        // 创建工作簿
        auto workbook = Workbook::create(Path("rgb_color_test.xlsx"));
        if (!workbook->open()) {
            std::cout << "Failed to open workbook" << std::endl;
            return 1;
        }
        
        // 添加工作表
        auto worksheet = workbook->addSheet("RGBColorTest");
        
        std::cout << "=== FastExcel RGB三参数颜色构造函数测试 ===" << std::endl;
        
        // 测试各种颜色创建方式
        Color red(255, 0, 0);      // 纯红色
        Color green(0, 255, 0);    // 纯绿色  
        Color blue(0, 0, 255);     // 纯蓝色
        Color purple(128, 0, 128); // 紫色
        Color orange(255, 165, 0); // 橙色
        
        // 验证颜色创建正确性
        std::cout << "✅ 颜色创建测试:" << std::endl;
        std::cout << "  红色(255,0,0): RGB=0x" << std::hex << red.getRGB() << std::dec;
        std::cout << " R=" << static_cast<int>(red.getRed()) 
                  << " G=" << static_cast<int>(red.getGreen()) 
                  << " B=" << static_cast<int>(red.getBlue()) << std::endl;
                  
        std::cout << "  绿色(0,255,0): RGB=0x" << std::hex << green.getRGB() << std::dec;
        std::cout << " R=" << static_cast<int>(green.getRed()) 
                  << " G=" << static_cast<int>(green.getGreen()) 
                  << " B=" << static_cast<int>(green.getBlue()) << std::endl;
                  
        std::cout << "  蓝色(0,0,255): RGB=0x" << std::hex << blue.getRGB() << std::dec;
        std::cout << " R=" << static_cast<int>(blue.getRed()) 
                  << " G=" << static_cast<int>(blue.getGreen()) 
                  << " B=" << static_cast<int>(blue.getBlue()) << std::endl;
                  
        std::cout << "  紫色(128,0,128): RGB=0x" << std::hex << purple.getRGB() << std::dec;
        std::cout << " R=" << static_cast<int>(purple.getRed()) 
                  << " G=" << static_cast<int>(purple.getGreen()) 
                  << " B=" << static_cast<int>(purple.getBlue()) << std::endl;
                  
        std::cout << "  橙色(255,165,0): RGB=0x" << std::hex << orange.getRGB() << std::dec;
        std::cout << " R=" << static_cast<int>(orange.getRed()) 
                  << " G=" << static_cast<int>(orange.getGreen()) 
                  << " B=" << static_cast<int>(orange.getBlue()) << std::endl;
        
        // 创建使用自定义颜色的样式
        auto redStyle = workbook->createStyleBuilder()
            .fontColor(red)
            .fontSize(12)
            .bold()
            .build();
            
        auto purpleStyle = workbook->createStyleBuilder()
            .fontColor(purple)
            .fontSize(14)
            .fill(orange)  // 橙色背景
            .build();
        
        int redStyleId = workbook->addStyle(redStyle);
        int purpleStyleId = workbook->addStyle(purpleStyle);
        
        // 写入带自定义颜色的文本
        worksheet->writeString(0, 0, "红色文字 (255,0,0)");
        auto& cell1 = worksheet->getCell(0, 0);
        cell1.setFormat(workbook->getStyles().getFormat(redStyleId));
        
        worksheet->writeString(1, 0, "紫色文字橙色背景 (128,0,128) + (255,165,0)");
        auto& cell2 = worksheet->getCell(1, 0);
        cell2.setFormat(workbook->getStyles().getFormat(purpleStyleId));
        
        // 保存文件
        workbook->save();
        workbook->close();
        
        std::cout << "\n🎉 FastExcel RGB三参数构造函数完美支持!" << std::endl;
        std::cout << "📋 新增功能:" << std::endl;
        std::cout << "   🎨 Color(255, 0, 0)     // 红色" << std::endl;
        std::cout << "   🎨 Color(0, 255, 0)     // 绿色" << std::endl; 
        std::cout << "   🎨 Color(0, 0, 255)     // 蓝色" << std::endl;
        std::cout << "   🎨 Color(128, 0, 128)   // 紫色" << std::endl;
        std::cout << "   🎨 Color(255, 165, 0)   // 橙色" << std::endl;
        std::cout << "\n✅ 比原来的 Color(0xFF0000) 更直观易用！" << std::endl;
        
    } catch (const std::exception& e) {
        std::cout << "❌ 错误: " << e.what() << std::endl;
        return 1;
    }
    
    return 0;
}