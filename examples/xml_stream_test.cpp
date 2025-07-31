#include "fastexcel/xml/XMLStreamWriter.hpp"
#include <iostream>
#include <fstream>
#include <chrono>

using namespace fastexcel::xml;

int main() {
    std::cout << "=== XMLStreamWriter 流式XML写入器测试 ===" << std::endl;
    
    // 测试1: 基本文档创建
    std::cout << "\n1. 测试基本文档创建..." << std::endl;
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
        std::ofstream out("test_basic.xml");
        out << result;
        out.close();
        std::cout << "已保存到 test_basic.xml" << std::endl;
    }
    
    // 测试2: 复杂嵌套结构
    std::cout << "\n2. 测试复杂嵌套结构..." << std::endl;
    {
        XMLStreamWriter writer;
        writer.startDocument();
        writer.startElement("bookstore");
        writer.writeAttribute("location", "New York");
        
        // 添加第一本书
        writer.startElement("book");
        writer.writeAttribute("category", "fiction");
        writer.writeAttribute("lang", "en");
        
        writer.startElement("title");
        writer.writeText("The Great Gatsby");
        writer.endElement();
        
        writer.startElement("author");
        writer.writeText("F. Scott Fitzgerald");
        writer.endElement();
        
        writer.startElement("year");
        writer.writeText("1925");
        writer.endElement();
        
        writer.startElement("price");
        writer.writeText("12.99");
        writer.endElement();
        
        writer.endElement(); // book
        
        // 添加第二本书
        writer.startElement("book");
        writer.writeAttribute("category", "science-fiction");
        writer.writeAttribute("lang", "en");
        
        writer.startElement("title");
        writer.writeText("Dune");
        writer.endElement();
        
        writer.startElement("author");
        writer.writeText("Frank Herbert");
        writer.endElement();
        
        writer.startElement("year");
        writer.writeText("1965");
        writer.endElement();
        
        writer.startElement("price");
        writer.writeText("8.99");
        writer.endElement();
        
        writer.endElement(); // book
        
        writer.endElement(); // bookstore
        writer.endDocument();
        
        std::string result = writer.toString();
        
        // 保存到文件
        std::ofstream out("test_complex.xml");
        out << result;
        out.close();
        std::cout << "已保存到 test_complex.xml" << std::endl;
    }
    
    // 测试3: 字符转义
    std::cout << "\n3. 测试字符转义..." << std::endl;
    {
        XMLStreamWriter writer;
        writer.startDocument();
        writer.startElement("test");
        writer.writeAttribute("attr", "Special chars: < > & \" ' \n");
        writer.writeText("Text with special chars: < > & \" ' \n");
        writer.endElement();
        writer.endDocument();
        
        std::string result = writer.toString();
        std::cout << "生成的XML: " << result << std::endl;
        
        // 保存到文件
        std::ofstream out("test_escape.xml");
        out << result;
        out.close();
        std::cout << "已保存到 test_escape.xml" << std::endl;
    }
    
    // 测试4: 直接文件写入模式
    std::cout << "\n4. 测试直接文件写入模式..." << std::endl;
    {
        FILE* file = nullptr;
        errno_t err = fopen_s(&file, "test_direct.xml", "wb");
        if (err != 0 || !file) {
            std::cerr << "无法创建文件" << std::endl;
            return 1;
        }
        
        XMLStreamWriter writer;
        writer.setDirectFileMode(file, true);  // 自动管理文件句柄
        
        writer.startDocument();
        writer.startElement("root");
        writer.writeAttribute("mode", "direct");
        
        for (int i = 0; i < 100; ++i) {
            writer.startElement("item");
            writer.writeAttribute("id", std::to_string(i));
            writer.writeText("Item " + std::to_string(i));
            writer.endElement();
        }
        
        writer.endElement();
        writer.endDocument();
        // 文件自动关闭
        
        std::cout << "已保存到 test_direct.xml" << std::endl;
    }
    
    // 测试5: 性能测试
    std::cout << "\n5. 性能测试..." << std::endl;
    {
        const int elements = 10000;
        
        auto start = std::chrono::high_resolution_clock::now();
        
        XMLStreamWriter writer;
        writer.startDocument();
        writer.startElement("root");
        
        for (int i = 0; i < elements; ++i) {
            writer.startElement("item");
            writer.writeAttribute("id", std::to_string(i));
            writer.writeAttribute("name", "item_" + std::to_string(i));
            writer.writeText("Content for item " + std::to_string(i));
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
        std::ofstream out("test_performance.xml");
        out << writer.toString();
        out.close();
        std::cout << "已保存到 test_performance.xml" << std::endl;
    }
    
    // 测试6: 大文件测试
    std::cout << "\n6. 大文件测试..." << std::endl;
    {
        FILE* file = nullptr;
        errno_t err = fopen_s(&file, "test_large.xml", "wb");
        if (err != 0 || !file) {
            std::cerr << "无法创建文件" << std::endl;
            return 1;
        }
        
        XMLStreamWriter writer;
        writer.setDirectFileMode(file, true);
        
        writer.startDocument();
        writer.startElement("root");
        writer.writeAttribute("description", "Large file test");
        
        const int large_elements = 50000;
        for (int i = 0; i < large_elements; ++i) {
            writer.startElement("item");
            writer.writeAttribute("id", std::to_string(i));
            writer.writeText("This is a longer text content for item " + std::to_string(i) + 
                           " to test the performance with larger text content.");
            writer.endElement();
            
            if (i % 10000 == 0) {
                std::cout << "已处理 " << i << " 个元素..." << std::endl;
            }
        }
        
        writer.endElement();
        writer.endDocument();
        
        std::cout << "已保存到 test_large.xml" << std::endl;
    }
    
    std::cout << "\n=== 所有测试完成 ===" << std::endl;
    std::cout << "请检查生成的XML文件以验证结果。" << std::endl;
    
    return 0;
}