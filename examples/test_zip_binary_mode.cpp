/**
 * @file test_zip_binary_mode.cpp
 * @brief 测试二进制模式和文本模式对ZIP文件的影响
 */

#include <iostream>
#include <fstream>
#include <sstream>
#include <vector>
#include <string>
#include <cstring>
#include "fastexcel/archive/ZipArchive.hpp"
#include "fastexcel/utils/Logger.hpp"

using namespace fastexcel::archive;

// 测试XML内容
const char* TEST_XML = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr">
        <is><t>测试内容</t></is>
      </c>
    </row>
  </sheetData>
</worksheet>)";

void hexDump(const std::string& label, const void* data, size_t size) {
    std::cout << "\n" << label << " (size=" << size << "):" << std::endl;
    const unsigned char* bytes = static_cast<const unsigned char*>(data);
    for (size_t i = 0; i < size && i < 64; ++i) {
        std::cout << std::hex << std::setw(2) << std::setfill('0') 
                  << static_cast<int>(bytes[i]) << " ";
        if ((i + 1) % 16 == 0) std::cout << std::endl;
    }
    if (size > 64) std::cout << "... (truncated)";
    std::cout << std::dec << std::endl;
}

void testDirectString() {
    std::cout << "\n=== 测试1: 直接字符串方式 ===" << std::endl;
    
    ZipArchive zip("test_direct_string.xlsx");
    if (!zip.open(true)) {
        std::cerr << "无法创建ZIP文件" << std::endl;
        return;
    }
    
    // 直接使用字符串
    std::string content = TEST_XML;
    std::cout << "String length: " << content.length() << std::endl;
    std::cout << "String size: " << content.size() << std::endl;
    hexDump("Direct string content", content.data(), content.size());
    
    auto result = zip.addFile("xl/worksheets/sheet1.xml", content);
    if (result != ZipError::Ok) {
        std::cerr << "添加文件失败" << std::endl;
    }
    
    zip.close();
    std::cout << "创建文件: test_direct_string.xlsx" << std::endl;
}

void testBinaryMode() {
    std::cout << "\n=== 测试2: 二进制模式读取 ===" << std::endl;
    
    // 先创建一个临时文件
    {
        std::ofstream temp("temp_test.xml", std::ios::binary);
        temp.write(TEST_XML, std::strlen(TEST_XML));
        temp.close();
    }
    
    // 以二进制模式读取
    std::ifstream file("temp_test.xml", std::ios::binary);
    std::string content((std::istreambuf_iterator<char>(file)),
                       std::istreambuf_iterator<char>());
    file.close();
    
    std::cout << "Binary read length: " << content.length() << std::endl;
    std::cout << "Binary read size: " << content.size() << std::endl;
    hexDump("Binary mode content", content.data(), content.size());
    
    ZipArchive zip("test_binary_mode.xlsx");
    if (!zip.open(true)) {
        std::cerr << "无法创建ZIP文件" << std::endl;
        return;
    }
    
    auto result = zip.addFile("xl/worksheets/sheet1.xml", content);
    if (result != ZipError::Ok) {
        std::cerr << "添加文件失败" << std::endl;
    }
    
    zip.close();
    std::cout << "创建文件: test_binary_mode.xlsx" << std::endl;
    
    // 清理临时文件
    std::remove("temp_test.xml");
}

void testVectorUint8() {
    std::cout << "\n=== 测试3: vector<uint8_t>方式 ===" << std::endl;
    
    ZipArchive zip("test_vector_uint8.xlsx");
    if (!zip.open(true)) {
        std::cerr << "无法创建ZIP文件" << std::endl;
        return;
    }
    
    // 使用vector<uint8_t>
    std::vector<uint8_t> data(TEST_XML, TEST_XML + std::strlen(TEST_XML));
    std::cout << "Vector size: " << data.size() << std::endl;
    hexDump("Vector<uint8_t> content", data.data(), data.size());
    
    auto result = zip.addFile("xl/worksheets/sheet1.xml", data.data(), data.size());
    if (result != ZipError::Ok) {
        std::cerr << "添加文件失败" << std::endl;
    }
    
    zip.close();
    std::cout << "创建文件: test_vector_uint8.xlsx" << std::endl;
}

void testWithBOM() {
    std::cout << "\n=== 测试4: 带BOM的UTF-8 ===" << std::endl;
    
    ZipArchive zip("test_with_bom.xlsx");
    if (!zip.open(true)) {
        std::cerr << "无法创建ZIP文件" << std::endl;
        return;
    }
    
    // 添加UTF-8 BOM
    std::string content;
    content.push_back(0xEF);
    content.push_back(0xBB);
    content.push_back(0xBF);
    content.append(TEST_XML);
    
    std::cout << "Content with BOM size: " << content.size() << std::endl;
    hexDump("Content with BOM", content.data(), content.size());
    
    auto result = zip.addFile("xl/worksheets/sheet1.xml", content);
    if (result != ZipError::Ok) {
        std::cerr << "添加文件失败" << std::endl;
    }
    
    zip.close();
    std::cout << "创建文件: test_with_bom.xlsx" << std::endl;
}

void testCompleteExcel() {
    std::cout << "\n=== 测试5: 完整的Excel文件结构 ===" << std::endl;
    
    ZipArchive zip("test_complete_excel.xlsx");
    if (!zip.open(true)) {
        std::cerr << "无法创建ZIP文件" << std::endl;
        return;
    }
    
    // 创建完整的Excel结构
    std::vector<ZipArchive::FileEntry> files;
    
    // [Content_Types].xml - 必须的根文件
    files.emplace_back("[Content_Types].xml", 
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>)");
    
    // _rels/.rels - 必须的关系文件
    files.emplace_back("_rels/.rels",
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>)");
    
    // xl/workbook.xml - 工作簿
    files.emplace_back("xl/workbook.xml",
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>)");
    
    // xl/_rels/workbook.xml.rels
    files.emplace_back("xl/_rels/workbook.xml.rels",
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>)");
    
    // xl/worksheets/sheet1.xml
    files.emplace_back("xl/worksheets/sheet1.xml", TEST_XML);
    
    // 批量添加所有文件
    auto result = zip.addFiles(files);
    if (result != ZipError::Ok) {
        std::cerr << "批量添加文件失败" << std::endl;
    }
    
    zip.close();
    std::cout << "创建文件: test_complete_excel.xlsx" << std::endl;
}

int main() {
    // 初始化日志系统
    fastexcel::Logger::getInstance().initialize("logs/zip_binary_mode_test.log",
                                               fastexcel::Logger::Level::DEBUG,
                                               true);
    
    std::cout << "ZIP二进制模式测试" << std::endl;
    std::cout << "==================" << std::endl;
    
    // 运行各种测试
    testDirectString();
    testBinaryMode();
    testVectorUint8();
    testWithBOM();
    testCompleteExcel();
    
    std::cout << "\n测试完成！" << std::endl;
    std::cout << "请使用Excel打开生成的文件，查看哪些能正常打开。" << std::endl;
    std::cout << "同时检查日志文件 logs/zip_binary_mode_test.log 查看详细信息。" << std::endl;
    
    // 关闭日志系统
    fastexcel::Logger::getInstance().shutdown();
    
    return 0;
}