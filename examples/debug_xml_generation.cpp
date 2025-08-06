/**
 * @file debug_xml_generation.cpp
 * @brief 调试XML生成问题
 */

#include "fastexcel/FastExcel.hpp"

#include <cassert>
#include <fstream>
#include <iostream>

using namespace fastexcel;
using namespace fastexcel::core;

int main() {
    try {
        // 初始化FastExcel
        void (*init_func)() = fastexcel::initialize;
        init_func();
        
        std::cout << "调试XML生成问题" << std::endl;
        std::cout << "==================" << std::endl;
        
        // 创建工作簿和工作表
        auto workbook = Workbook::create("debug_xml.xlsx");
        assert(workbook->open());
        
        // 禁用共享字符串表以便在XML中直接显示字符串
        workbook->setUseSharedStrings(false);
        
        auto worksheet = workbook->addWorksheet("TestSheet");
        
        // 写入测试数据
        std::cout << "写入数据到单元格 (0,0): 'Hello'" << std::endl;
        worksheet->writeString(0, 0, "Hello");
        
        std::cout << "写入数据到单元格 (0,1): 123.45" << std::endl;
        worksheet->writeNumber(0, 1, 123.45);
        
        // 验证数据写入
        auto& cell = worksheet->getCell(0, 0);
        std::cout << "单元格 (0,0) 是否为空: " << (cell.isEmpty() ? "是" : "否") << std::endl;
        if (!cell.isEmpty()) {
            std::cout << "单元格 (0,0) 类型: " << (cell.isString() ? "字符串" : "其他") << std::endl;
            if (cell.isString()) {
                std::cout << "单元格 (0,0) 值: '" << cell.getStringValue() << "'" << std::endl;
            }
        }
        
        // 获取使用范围
        auto [max_row, max_col] = worksheet->getUsedRange();
        std::cout << "使用范围: 行=" << max_row << ", 列=" << max_col << std::endl;
        
        // 生成XML
        std::cout << "\n生成XML..." << std::endl;
        std::string xml;
        worksheet->generateXML([&xml](const char* data, size_t size) {
            xml.append(data, size);
        });
        
        std::cout << "XML长度: " << xml.length() << " 字符" << std::endl;
        
        if (xml.empty()) {
            std::cout << "❌ XML为空！" << std::endl;
        } else {
            // 保存完整XML到文件
            std::ofstream file("debug_output.xml");
            file << xml;
            file.close();
            std::cout << "✓ XML已保存到 debug_output.xml" << std::endl;
            
            // 检查是否包含"Hello"
            if (xml.find("Hello") != std::string::npos) {
                std::cout << "✓ XML中找到了'Hello'" << std::endl;
            } else {
                std::cout << "❌ XML中未找到'Hello'" << std::endl;
                
                // 输出XML的前1000个字符用于调试
                std::cout << "\nXML预览 (前1000字符):" << std::endl;
                std::cout << "========================================" << std::endl;
                std::cout << xml.substr(0, 1000) << std::endl;
                std::cout << "========================================" << std::endl;
            }
            
            // 检查其他关键元素
            if (xml.find("<worksheet") != std::string::npos) {
                std::cout << "✓ 找到了<worksheet元素" << std::endl;
            } else {
                std::cout << "❌ 未找到<worksheet元素" << std::endl;
            }
            
            if (xml.find("<sheetData") != std::string::npos) {
                std::cout << "✓ 找到了<sheetData元素" << std::endl;
            } else {
                std::cout << "❌ 未找到<sheetData元素" << std::endl;
            }
        }
        
        workbook->close();
        fastexcel::cleanup();
        
    } catch (const std::exception& e) {
        std::cerr << "错误: " << e.what() << std::endl;
        return 1;
    }
    
    return 0;
}
