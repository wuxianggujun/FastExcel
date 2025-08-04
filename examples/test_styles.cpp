#include "../src/fastexcel/core/Format.hpp"
#include "../src/fastexcel/core/FormatPool.hpp"
#include <fstream>
#include <iostream>

int main() {
    try {
        std::cout << "Testing styles XML generation..." << std::endl;
        
        // 创建格式池
        fastexcel::core::FormatPool pool;
        
        // 创建红色字体 + 黄色背景的格式
        auto format1 = std::make_unique<fastexcel::core::Format>();
        format1->setFontColor(fastexcel::core::Color(0xFF0000U)); // 红色字体
        format1->setBackgroundColor(fastexcel::core::Color(0xFFFF00U)); // 黄色背景
        
        // 创建粗体 + 居中对齐的格式
        auto format2 = std::make_unique<fastexcel::core::Format>();
        format2->setBold(true);
        format2->setHorizontalAlign(fastexcel::core::HorizontalAlign::Center);
        
        // 添加格式到池中
        pool.addFormat(std::move(format1));
        pool.addFormat(std::move(format2));
        
        // 生成样式XML到文件
        pool.generateStylesXMLToFile("test_styles.xml");
        
        std::cout << "Styles XML generated successfully to test_styles.xml" << std::endl;
        
        // 读取并显示生成的XML内容
        std::ifstream file("test_styles.xml");
        if (file.is_open()) {
            std::cout << "\nGenerated XML content:" << std::endl;
            std::cout << "========================" << std::endl;
            std::string line;
            while (std::getline(file, line)) {
                std::cout << line << std::endl;
            }
            file.close();
        }
        
        return 0;
        
    } catch (const std::exception& e) {
        std::cerr << "Exception: " << e.what() << std::endl;
        return 1;
    }
}