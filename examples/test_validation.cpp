/**
 * @file test_validation.cpp
 * @brief FastExcel功能验证示例程序
 * 
 * 这个程序用来验证FastExcel的核心功能是否正常工作，
 * 包括工作簿创建、工作表操作、格式设置、XML生成等
 */

#include "fastexcel/FastExcel.hpp"
#include <iostream>
#include <cassert>
#include <string>

using namespace fastexcel;
using namespace fastexcel::core;

void testBasicWorkbookOperations() {
    std::cout << "=== 测试基本工作簿操作 ===" << std::endl;
    
    // 初始化FastExcel
    void (*init_func)() = fastexcel::initialize;
    init_func();
    
    // 创建工作簿
    auto workbook = Workbook::create("test_validation.xlsx");
    assert(workbook != nullptr);
    std::cout << "✓ 工作簿创建成功" << std::endl;
    
    // 打开工作簿
    assert(workbook->open());
    std::cout << "✓ 工作簿打开成功" << std::endl;
    
    // 添加工作表
    auto worksheet = workbook->addWorksheet("TestSheet");
    assert(worksheet != nullptr);
    std::cout << "✓ 工作表创建成功: " << worksheet->getName() << std::endl;
    
    // 测试单元格写入
    worksheet->writeString(0, 0, "Hello");
    worksheet->writeNumber(0, 1, 123.45);
    worksheet->writeBoolean(0, 2, true);
    std::cout << "✓ 单元格数据写入成功" << std::endl;
    
    // 验证单元格数据读取
    auto& hello_cell = worksheet->getCell(0, 0);
    assert(hello_cell.isString());
    assert(hello_cell.getStringValue() == "Hello");
    
    auto& number_cell = worksheet->getCell(0, 1);
    assert(number_cell.isNumber());
    assert(std::abs(number_cell.getNumberValue() - 123.45) < 0.001);
    
    auto& bool_cell = worksheet->getCell(0, 2);
    assert(bool_cell.isBoolean());
    assert(bool_cell.getBooleanValue() == true);
    std::cout << "✓ 单元格数据读取验证成功" << std::endl;
    
    // 保存并关闭
    assert(workbook->save());
    std::cout << "✓ 工作簿保存成功" << std::endl;
    
    workbook->close();
    std::cout << "✓ 工作簿关闭成功" << std::endl;
    
    // 清理FastExcel
    fastexcel::cleanup();
}

void testFormatOperations() {
    std::cout << "\n=== 测试格式操作 ===" << std::endl;
    
    // 初始化FastExcel
    void (*init_func)() = fastexcel::initialize;
    init_func();
    
    auto workbook = Workbook::create("test_formats.xlsx");
    assert(workbook->open());
    
    auto worksheet = workbook->addWorksheet("FormatsSheet");
    
    // 创建不同的格式
    auto bold_format = workbook->createFormat();
    bold_format->setBold(true);
    
    auto italic_format = workbook->createFormat();
    italic_format->setItalic(true);
    
    auto colored_format = workbook->createFormat();
    colored_format->setFontColor(Color::red());
    colored_format->setFontSize(14);
    
    std::cout << "✓ 格式创建成功" << std::endl;
    
    // 应用格式写入数据
    worksheet->writeString(0, 0, "Bold Text", bold_format);
    worksheet->writeString(1, 0, "Italic Text", italic_format);
    worksheet->writeString(2, 0, "Colored Text", colored_format);
    
    std::cout << "✓ 格式化文本写入成功" << std::endl;
    
    // 验证格式池
    size_t format_count = workbook->getFormatCount();
    std::cout << "格式池中的格式数量: " << format_count << std::endl;
    assert(format_count >= 3); // 至少应该有我们创建的格式
    
    // 保存文件
    assert(workbook->save());
    workbook->close();
    
    std::cout << "✓ 格式测试完成" << std::endl;
    
    fastexcel::cleanup();
}

void testXMLGeneration() {
    std::cout << "\n=== 测试XML生成 ===" << std::endl;
    
    void (*init_func)() = fastexcel::initialize;
    init_func();
    
    auto workbook = Workbook::create("test_xml.xlsx");
    assert(workbook->open());
    
    auto worksheet = workbook->addWorksheet("XMLTestSheet");
    
    // 写入测试数据
    worksheet->writeString(0, 0, "XML Test");
    worksheet->writeNumber(0, 1, 42.0);
    
    // 生成XML并验证
    std::string xml;
    worksheet->generateXML([&xml](const char* data, size_t size) {
        xml.append(data, size);
    });
    
    std::cout << "生成的XML长度: " << xml.length() << " 字符" << std::endl;
    
    // 验证XML包含必要元素
    assert(!xml.empty());
    assert(xml.find("<worksheet") != std::string::npos);
    assert(xml.find("<sheetData") != std::string::npos);
    std::cout << "✓ XML基本结构验证成功" << std::endl;
    
    // 输出XML预览（前500个字符）
    std::cout << "XML预览:\n" << xml.substr(0, 500) << "..." << std::endl;
    
    workbook->close();
    fastexcel::cleanup();
    
    std::cout << "✓ XML生成测试完成" << std::endl;
}

void testWorksheetOperations() {
    std::cout << "\n=== 测试工作表操作 ===" << std::endl;
    
    void (*init_func)() = fastexcel::initialize;
    init_func();
    
    auto workbook = Workbook::create("test_worksheet_ops.xlsx");
    assert(workbook->open());
    
    auto worksheet = workbook->addWorksheet("OpsTest");
    
    // 测试批量数据写入
    std::vector<std::vector<std::string>> string_data = {
        {"A1", "B1", "C1"},
        {"A2", "B2", "C2"},
        {"A3", "B3", "C3"}
    };
    worksheet->writeRange(0, 0, string_data);
    std::cout << "✓ 批量字符串数据写入成功" << std::endl;
    
    std::vector<std::vector<double>> number_data = {
        {1.1, 2.2, 3.3},
        {4.4, 5.5, 6.6}
    };
    worksheet->writeRange(5, 0, number_data);
    std::cout << "✓ 批量数字数据写入成功" << std::endl;
    
    // 测试合并单元格
    worksheet->mergeCells(10, 0, 10, 2);
    worksheet->writeString(10, 0, "Merged Cell");
    std::cout << "✓ 合并单元格操作成功" << std::endl;
    
    // 测试列宽设置
    worksheet->setColumnWidth(0, 15.0);
    worksheet->setColumnWidth(1, 20.0);
    std::cout << "✓ 列宽设置成功" << std::endl;
    
    // 测试使用范围获取
    auto [max_row, max_col] = worksheet->getUsedRange();
    std::cout << "使用范围: 行 " << max_row << ", 列 " << max_col << std::endl;
    
    assert(workbook->save());
    workbook->close();
    fastexcel::cleanup();
    
    std::cout << "✓ 工作表操作测试完成" << std::endl;
}

void testMemoryAndPerformance() {
    std::cout << "\n=== 测试内存和性能 ===" << std::endl;
    
    void (*init_func)() = fastexcel::initialize;
    init_func();
    
    auto workbook = Workbook::create("test_performance.xlsx");
    assert(workbook->open());
    
    auto worksheet = workbook->addWorksheet("PerfTest");
    
    // 写入大量数据测试性能
    const int rows = 1000;
    const int cols = 10;
    
    auto start_time = std::chrono::high_resolution_clock::now();
    
    for (int r = 0; r < rows; ++r) {
        for (int c = 0; c < cols; ++c) {
            if (c % 2 == 0) {
                worksheet->writeString(r, c, "Row" + std::to_string(r) + "Col" + std::to_string(c));
            } else {
                worksheet->writeNumber(r, c, r * cols + c);
            }
        }
    }
    
    auto end_time = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time);
    
    std::cout << "写入 " << (rows * cols) << " 个单元格耗时: " << duration.count() << " 毫秒" << std::endl;
    
    assert(workbook->save());
    workbook->close();
    fastexcel::cleanup();
    
    std::cout << "✓ 性能测试完成" << std::endl;
}

int main() {
    std::cout << "FastExcel 功能验证程序" << std::endl;
    std::cout << "版本: " << fastexcel::getVersion() << std::endl;
    std::cout << "========================================" << std::endl;
    
    try {
        testBasicWorkbookOperations();
        testFormatOperations();
        testXMLGeneration();
        testWorksheetOperations();
        testMemoryAndPerformance();
        
        std::cout << "\n========================================" << std::endl;
        std::cout << "🎉 所有测试通过！FastExcel功能正常。" << std::endl;
        
    } catch (const std::exception& e) {
        std::cerr << "\n❌ 测试失败: " << e.what() << std::endl;
        return 1;
    } catch (...) {
        std::cerr << "\n❌ 未知错误" << std::endl;
        return 1;
    }
    
    return 0;
}