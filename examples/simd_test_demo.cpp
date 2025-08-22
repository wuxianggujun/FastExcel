#include "fastexcel/xml/XMLEscapeSIMD.hpp"
#include "fastexcel/xml/XMLStreamWriter.hpp"
#include <iostream>
#include <chrono>
#include <string>

using namespace fastexcel::xml;

int main() {
    std::cout << "=== FastExcel SIMD优化测试 ===" << std::endl;
    
    // 初始化SIMD支持
    XMLEscapeSIMD::initialize();
    
    // 检查SIMD支持状态
    bool simd_supported = XMLEscapeSIMD::isSIMDSupported();
    std::cout << "SIMD支持状态: " << (simd_supported ? "支持" : "不支持") << std::endl;
    
    // 测试数据 - 包含需要转义的字符
    std::string test_data = "这是一个测试 & 包含 < 特殊字符 > 和 \"引号\" 以及 '单引号' 还有换行符\n的XML数据";
    std::cout << "测试数据长度: " << test_data.length() << " 字符" << std::endl;
    
    // 测试XMLStreamWriter的SIMD集成
    std::string result;
    XMLStreamWriter writer([&result](const char* data, size_t len) {
        result.append(data, len);
    });
    
    // 开始计时
    auto start = std::chrono::high_resolution_clock::now();
    
    // 执行多次转义操作来测试性能
    const int iterations = 10000;
    for (int i = 0; i < iterations; i++) {
        writer.startDocument();
        writer.startElement("test");
        writer.writeAttribute("attr", test_data.c_str());
        writer.writeText(test_data.c_str());
        writer.endElement();
        writer.endDocument();
    }
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    std::cout << "执行 " << iterations << " 次转义操作耗时: " 
              << duration.count() << " 微秒" << std::endl;
    std::cout << "平均每次操作: " << (double)duration.count() / iterations << " 微秒" << std::endl;
    
    // 验证转义是否正确
    if (result.find("&amp;") != std::string::npos && 
        result.find("&lt;") != std::string::npos && 
        result.find("&gt;") != std::string::npos) {
        std::cout << "✓ XML转义功能正常工作" << std::endl;
    } else {
        std::cout << "✗ XML转义功能异常" << std::endl;
    }
    
    // 测试原始SIMD功能
    std::cout << "\n--- 直接SIMD转义测试 ---" << std::endl;
    std::string simd_result;
    
    auto simd_start = std::chrono::high_resolution_clock::now();
    
    for (int i = 0; i < iterations; i++) {
        simd_result.clear();
        XMLEscapeSIMD::escapeAttributesSIMD(test_data.c_str(), test_data.length(),
            [&simd_result](const char* data, size_t len) {
                simd_result.append(data, len);
            });
    }
    
    auto simd_end = std::chrono::high_resolution_clock::now();
    auto simd_duration = std::chrono::duration_cast<std::chrono::microseconds>(simd_end - simd_start);
    
    std::cout << "直接SIMD转义 " << iterations << " 次耗时: " 
              << simd_duration.count() << " 微秒" << std::endl;
    std::cout << "平均每次操作: " << (double)simd_duration.count() / iterations << " 微秒" << std::endl;
    
    std::cout << "\n✓ SIMD优化集成测试完成！" << std::endl;
    
    return 0;
}