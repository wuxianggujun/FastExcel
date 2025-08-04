#include "../src/fastexcel/core/Format.hpp"
#include "../src/fastexcel/core/Workbook.hpp"
#include <iostream>

int main() {
    try {
        std::cout << "Creating simple format test..." << std::endl;
        
        // 创建工作簿
        auto workbook = fastexcel::core::Workbook::create("simple_format_test.xlsx");
        if (!workbook->open()) {
            std::cerr << "Failed to open workbook" << std::endl;
            return 1;
        }

        // 添加工作表
        auto worksheet = workbook->addWorksheet("Test");
        if (!worksheet) {
            std::cerr << "Failed to add worksheet" << std::endl;
            return 1;
        }

        // 创建一个简单的格式：红色字体 + 黄色背景
        auto format1 = workbook->createFormat();
        format1->setFontColor(fastexcel::core::Color(0xFF0000U)); // 红色字体
        format1->setBackgroundColor(fastexcel::core::Color(0xFFFF00U)); // 黄色背景
        
        // 创建另一个格式：粗体 + 居中对齐
        auto format2 = workbook->createFormat();
        format2->setBold(true);
        format2->setHorizontalAlign(fastexcel::core::HorizontalAlign::Center);

        // 写入数据
        worksheet->writeString(0, 0, "Red text on yellow", format1);
        worksheet->writeString(1, 0, "Bold centered text", format2);
        worksheet->writeString(2, 0, "Normal text");

        // 保存文件
        if (!workbook->save()) {
            std::cerr << "Failed to save workbook" << std::endl;
            return 1;
        }
        
        workbook->close();
        std::cout << "Simple format test file created successfully!" << std::endl;
        return 0;
        
    } catch (const std::exception& e) {
        std::cerr << "Exception: " << e.what() << std::endl;
        return 1;
    }
}
