/**
 * @file test_reader.cpp
 * @brief XLSXReader单元测试
 */

#include <gtest/gtest.h>
#include "fastexcel/reader/XLSXReader.hpp"
#include "fastexcel/reader/SharedStringsParser.hpp"
#include "fastexcel/reader/WorksheetParser.hpp"
#include "fastexcel/core/Workbook.hpp"
#include <fstream>
#include <sstream>
#include <chrono>

using namespace fastexcel::reader;
using namespace fastexcel::core;

class XLSXReaderTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 测试前的设置
    }
    
    void TearDown() override {
        // 测试后的清理
    }
};

// 测试SharedStringsParser
TEST_F(XLSXReaderTest, SharedStringsParserBasic) {
    SharedStringsParser parser;
    
    // 测试空内容
    EXPECT_TRUE(parser.parse(""));
    EXPECT_EQ(parser.getStringCount(), 0);
    
    // 测试简单的共享字符串XML
    std::string xml = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
    <si><t>Hello</t></si>
    <si><t>World</t></si>
    <si><t>Test</t></si>
</sst>)";
    
    EXPECT_TRUE(parser.parse(xml));
    EXPECT_EQ(parser.getStringCount(), 3);
    EXPECT_EQ(parser.getString(0), "Hello");
    EXPECT_EQ(parser.getString(1), "World");
    EXPECT_EQ(parser.getString(2), "Test");
    EXPECT_EQ(parser.getString(3), ""); // 超出范围
}

TEST_F(XLSXReaderTest, SharedStringsParserWithEntities) {
    SharedStringsParser parser;
    
    // 测试包含XML实体的字符串
    std::string xml = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
    <si><t>&lt;tag&gt;</t></si>
    <si><t>A &amp; B</t></si>
</sst>)";
    
    EXPECT_TRUE(parser.parse(xml));
    EXPECT_EQ(parser.getStringCount(), 2);
    EXPECT_EQ(parser.getString(0), "<tag>");
    EXPECT_EQ(parser.getString(1), "A & B");
}

TEST_F(XLSXReaderTest, SharedStringsParserRichText) {
    SharedStringsParser parser;
    
    // 测试富文本格式
    std::string xml = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
    <si>
        <r><t>Bold</t></r>
        <r><t> and </t></r>
        <r><t>Italic</t></r>
    </si>
</sst>)";
    
    EXPECT_TRUE(parser.parse(xml));
    EXPECT_EQ(parser.getStringCount(), 1);
    EXPECT_EQ(parser.getString(0), "Bold and Italic");
}

// 测试WorksheetParser的辅助方法
TEST_F(XLSXReaderTest, WorksheetParserCellReference) {
    WorksheetParser parser;
    
    // 使用反射或友元类来测试私有方法，这里我们创建一个测试用的公共包装
    // 由于parseCellReference是私有方法，我们需要通过其他方式测试
    // 这里先跳过，在实际项目中可以考虑将这些方法设为protected或添加测试友元
}

// 测试XLSXReader基本功能
TEST_F(XLSXReaderTest, XLSXReaderConstruction) {
    XLSXReader reader("test.xlsx");
    
    // 测试构造函数不会抛出异常
    EXPECT_NO_THROW(XLSXReader("another_test.xlsx"));
}

TEST_F(XLSXReaderTest, XLSXReaderOpenNonExistentFile) {
    XLSXReader reader("non_existent_file.xlsx");
    
    // 测试打开不存在的文件
    EXPECT_FALSE(reader.open());
}

// 测试XML实体解码
TEST_F(XLSXReaderTest, XMLEntityDecoding) {
    SharedStringsParser parser;
    
    // 测试各种XML实体
    std::string xml = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="5" uniqueCount="5">
    <si><t>&lt;</t></si>
    <si><t>&gt;</t></si>
    <si><t>&amp;</t></si>
    <si><t>&quot;</t></si>
    <si><t>&apos;</t></si>
</sst>)";
    
    EXPECT_TRUE(parser.parse(xml));
    EXPECT_EQ(parser.getStringCount(), 5);
    EXPECT_EQ(parser.getString(0), "<");
    EXPECT_EQ(parser.getString(1), ">");
    EXPECT_EQ(parser.getString(2), "&");
    EXPECT_EQ(parser.getString(3), "\"");
    EXPECT_EQ(parser.getString(4), "'");
}

// 测试空的共享字符串项
TEST_F(XLSXReaderTest, EmptySharedStrings) {
    SharedStringsParser parser;
    
    std::string xml = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
    <si><t>First</t></si>
    <si/>
    <si><t>Third</t></si>
</sst>)";
    
    EXPECT_TRUE(parser.parse(xml));
    EXPECT_EQ(parser.getStringCount(), 3);
    EXPECT_EQ(parser.getString(0), "First");
    EXPECT_EQ(parser.getString(1), "");
    EXPECT_EQ(parser.getString(2), "Third");
}

// 性能测试 - 大量共享字符串
TEST_F(XLSXReaderTest, LargeSharedStringsPerformance) {
    SharedStringsParser parser;
    
    // 生成包含1000个字符串的XML
    std::ostringstream xml_stream;
    xml_stream << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1000" uniqueCount="1000">)";
    
    for (int i = 0; i < 1000; ++i) {
        xml_stream << "<si><t>String" << i << "</t></si>";
    }
    
    xml_stream << "</sst>";
    
    std::string xml = xml_stream.str();
    
    // 测试解析性能
    auto start = std::chrono::high_resolution_clock::now();
    EXPECT_TRUE(parser.parse(xml));
    auto end = std::chrono::high_resolution_clock::now();
    
    auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);
    std::cout << "解析1000个共享字符串耗时: " << duration.count() << "ms" << std::endl;
    
    EXPECT_EQ(parser.getStringCount(), 1000);
    EXPECT_EQ(parser.getString(0), "String0");
    EXPECT_EQ(parser.getString(999), "String999");
}

// 测试错误处理
TEST_F(XLSXReaderTest, MalformedXML) {
    SharedStringsParser parser;
    
    // 测试格式错误的XML
    std::string bad_xml = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
    <si><t>Unclosed tag</si>
</sst>)";
    
    // 应该能够处理格式错误的XML而不崩溃
    EXPECT_NO_THROW(parser.parse(bad_xml));
}

// 集成测试 - 如果有测试文件的话
TEST_F(XLSXReaderTest, DISABLED_IntegrationTest) {
    // 这个测试需要一个真实的XLSX文件
    // 在实际环境中，可以创建一个小的测试XLSX文件
    
    XLSXReader reader("test_data/sample.xlsx");
    
    if (reader.open()) {
        auto names = reader.getWorksheetNames();
        EXPECT_GT(names.size(), 0);
        
        if (!names.empty()) {
            auto worksheet = reader.loadWorksheet(names[0]);
            EXPECT_NE(worksheet, nullptr);
        }
        
        reader.close();
    }
}

int main(int argc, char** argv) {
    ::testing::InitGoogleTest(&argc, argv);
    return RUN_ALL_TESTS();
}