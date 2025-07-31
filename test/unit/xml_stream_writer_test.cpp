#include "fastexcel/xml/XMLStreamWriter.hpp"
#include <iostream>
#include <chrono>
#include <cstdio>
#include <string>

using namespace fastexcel::xml;

int main() {
    std::cout << "=== XMLStreamWriter 测试程序 ===" << std::endl;
    
    // 测试1: 基本功能测试
    std::cout << "\n1. 基本功能测试..." << std::endl;
    {
        XMLStreamWriter writer;
        writer.startDocument();
        writer.startElement("root");
        writer.writeAttribute("version", "1.0");
        writer.writeText("Hello World");
        writer.endElement();
        writer.endDocument();
        
        std::string result = writer.toString();
        std::cout << "生成的XML: " << result << std::endl;
        
        // 保存到文件
        FILE* file = fopen("test_basic.xml", "wb");
        if (file) {
            fwrite(result.c_str(), 1, result.length(), file);
            fclose(file);
            std::cout << "已保存到 test_basic.xml" << std::endl;
        }
    }
    
    // 测试2: 字符转义测试
    std::cout << "\n2. 字符转义测试..." << std::endl;
    {
        XMLStreamWriter writer;
        writer.startDocument();
        writer.startElement("test");
        writer.writeAttribute("attr", "Special: < > & \" ' \n");
        writer.writeText("Text with special chars: < > & \" ' \n");
        writer.endElement();
        writer.endDocument();
        
        std::string result = writer.toString();
        std::cout << "生成的XML: " << result << std::endl;
        
        // 保存到文件
        FILE* file = fopen("test_escape.xml", "wb");
        if (file) {
            fwrite(result.c_str(), 1, result.length(), file);
            fclose(file);
            std::cout << "已保存到 test_escape.xml" << std::endl;
        }
    }
    
    // 测试3: 大文件测试
    std::cout << "\n3. 大文件测试..." << std::endl;
    {
        FILE* file = fopen("test_large.xml", "wb");
        if (!file) {
            std::cerr << "无法创建文件" << std::endl;
            return 1;
        }
        
        XMLStreamWriter writer;
        writer.setDirectFileMode(file, true);
        
        writer.startDocument();
        writer.startElement("root");
        writer.writeAttribute("description", "Large file test");
        
        const int large_elements = 5000;
        for (int i = 0; i < large_elements; ++i) {
            writer.startElement("item");
            writer.writeAttribute("id", std::to_string(i).c_str());
            writer.writeText(("This is a longer text content for item " + std::to_string(i) + 
                           " to test the performance with larger text content.").c_str());
            writer.endElement();
            
            if (i % 1000 == 0) {
                std::cout << "已处理 " << i << " 个元素..." << std::endl;
            }
        }
        
        writer.endElement();
        writer.endDocument();
        
        std::cout << "已保存到 test_large.xml" << std::endl;
    }
    
    // 测试4: 性能测试
    std::cout << "\n4. 性能测试..." << std::endl;
    {
        const int elements = 10000;
        
        auto start = std::chrono::high_resolution_clock::now();
        
        XMLStreamWriter writer;
        writer.startDocument();
        writer.startElement("root");
        
        for (int i = 0; i < elements; ++i) {
            writer.startElement("item");
            writer.writeAttribute("id", std::to_string(i).c_str());
            writer.writeAttribute("name", ("item_" + std::to_string(i)).c_str());
            writer.writeText(("Content for item " + std::to_string(i)).c_str());
            writer.endElement();
        }
        
        writer.endElement();
        writer.endDocument();
        
        auto end = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);
        
        std::cout << "生成 " << elements << " 个元素耗时: " 
                  << duration.count() << " 毫秒" << std::endl;
        std::cout << "平均每个元素: " 
                  << static_cast<double>(duration.count()) / elements << " 毫秒" << std::endl;
        
        // 保存到文件
        FILE* file = fopen("test_performance.xml", "wb");
        if (file) {
            std::string result = writer.toString();
            fwrite(result.c_str(), 1, result.length(), file);
            fclose(file);
            std::cout << "已保存到 test_performance.xml" << std::endl;
        }
    }
    
    // 测试5: 属性批处理测试
    std::cout << "\n5. 属性批处理测试..." << std::endl;
    {
        XMLStreamWriter writer;
        writer.startDocument();
        writer.startElement("product");
        
        // 使用批处理写入多个属性
        writer.startAttributeBatch();
        writer.writeAttribute("id", "12345");
        writer.writeAttribute("name", "Test Product");
        writer.writeAttribute("price", "99.99");
        writer.writeAttribute("category", "Electronics");
        writer.endAttributeBatch();
        
        writer.writeText("This is a test product with multiple attributes");
        writer.endElement();
        writer.endDocument();
        
        std::string result = writer.toString();
        std::cout << "生成的XML: " << result << std::endl;
        
        // 保存到文件
        FILE* file = fopen("test_batch.xml", "wb");
        if (file) {
            fwrite(result.c_str(), 1, result.length(), file);
            fclose(file);
            std::cout << "已保存到 test_batch.xml" << std::endl;
        }
    }
    
    std::cout << "\n=== 所有测试完成 ===" << std::endl;
    std::cout << "请检查生成的XML文件以验证结果。" << std::endl;
    
    return 0;
}