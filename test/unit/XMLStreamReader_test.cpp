#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/xml/XMLStreamReader.hpp"

#include <chrono>
#include <fstream>
#include <gtest/gtest.h>
#include <sstream>

namespace fastexcel {
namespace xml {

class XMLStreamReaderTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 初始化日志系统
        fastexcel::Logger::getInstance().initialize("logs/XMLStreamReader_test.log", 
                                                   fastexcel::Logger::Level::DEBUG, 
                                                   false);
        
        reader_ = std::make_unique<XMLStreamReader>();
    }

    void TearDown() override {
        reader_.reset();
        fastexcel::Logger::getInstance().shutdown();
    }

    std::unique_ptr<XMLStreamReader> reader_;
    
    // 测试用的XML内容
    const std::string simple_xml_ = R"(<?xml version="1.0" encoding="UTF-8"?>
<root>
    <element attr="value">Text content</element>
    <empty_element/>
    <parent>
        <child>Child text</child>
        <child>Another child</child>
    </parent>
</root>)";

    const std::string complex_xml_ = R"(<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheets>
        <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
        <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
    </sheets>
    <definedNames>
        <definedName name="Print_Area" localSheetId="0">'Sheet1'!$A$1:$C$10</definedName>
    </definedNames>
</workbook>)";
};

// 测试1: 基本解析功能
TEST_F(XMLStreamReaderTest, BasicParsing) {
    LOG_INFO("Testing basic XML parsing");

    std::vector<std::string> elements;
    std::vector<std::string> texts;
    
    reader_->setStartElementCallback([&](const std::string& name, const std::vector<XMLAttribute>& attributes, int depth) {
        elements.push_back(name);
        LOG_DEBUG("Start element: {} at depth {}", name, depth);
    });
    
    reader_->setTextCallback([&](const std::string& text, int depth) {
        if (!text.empty()) {
            texts.push_back(text);
            LOG_DEBUG("Text content: '{}' at depth {}", text, depth);
        }
    });
    
    XMLParseError result = reader_->parseFromString(simple_xml_);
    EXPECT_EQ(result, XMLParseError::Ok);
    
    // 验证解析的元素
    EXPECT_GE(elements.size(), 4);  // 至少应该有 root, element, empty_element, parent, child
    EXPECT_EQ(elements[0], "root");
    
    // 验证文本内容
    EXPECT_GE(texts.size(), 1);
    bool found_text_content = false;
    for (const auto& text : texts) {
        if (text == "Text content") {
            found_text_content = true;
            break;
        }
    }
    EXPECT_TRUE(found_text_content);
}

// 测试2: 属性解析
TEST_F(XMLStreamReaderTest, AttributeParsing) {
    LOG_INFO("Testing XML attribute parsing");

    std::unordered_map<std::string, std::string> found_attributes;
    
    reader_->setStartElementCallback([&](const std::string& name, const std::vector<XMLAttribute>& attributes, int depth) {
        for (const auto& attr : attributes) {
            found_attributes[attr.name] = attr.value;
            LOG_DEBUG("Attribute: {}='{}' in element '{}'", attr.name, attr.value, name);
        }
    });
    
    XMLParseError result = reader_->parseFromString(simple_xml_);
    EXPECT_EQ(result, XMLParseError::Ok);
    
    // 验证属性
    EXPECT_EQ(found_attributes["attr"], "value");
}

// 测试3: 复杂XML解析（模拟Excel格式）
TEST_F(XMLStreamReaderTest, ComplexXMLParsing) {
    LOG_INFO("Testing complex XML parsing (Excel-like format)");

    std::vector<std::pair<std::string, std::string>> sheets;
    
    reader_->setStartElementCallback([&](const std::string& name, const std::vector<XMLAttribute>& attributes, int depth) {
        if (name == "sheet") {
            std::string sheet_name, sheet_id;
            for (const auto& attr : attributes) {
                if (attr.name == "name") {
                    sheet_name = attr.value;
                } else if (attr.name == "sheetId") {
                    sheet_id = attr.value;
                }
            }
            if (!sheet_name.empty() && !sheet_id.empty()) {
                sheets.emplace_back(sheet_name, sheet_id);
                LOG_DEBUG("Found sheet: {} with ID {}", sheet_name, sheet_id);
            }
        }
    });
    
    XMLParseError result = reader_->parseFromString(complex_xml_);
    EXPECT_EQ(result, XMLParseError::Ok);
    
    // 验证解析的工作表
    EXPECT_EQ(sheets.size(), 2);
    EXPECT_EQ(sheets[0].first, "Sheet1");
    EXPECT_EQ(sheets[0].second, "1");
    EXPECT_EQ(sheets[1].first, "Sheet2");
    EXPECT_EQ(sheets[1].second, "2");
}

// 测试4: 流式解析
TEST_F(XMLStreamReaderTest, StreamParsing) {
    LOG_INFO("Testing stream-based XML parsing");

    std::vector<std::string> elements;
    
    reader_->setStartElementCallback([&](const std::string& name, const std::vector<XMLAttribute>& attributes, int depth) {
        elements.push_back(name);
    });
    
    // 开始流式解析
    XMLParseError result = reader_->beginParsing();
    EXPECT_EQ(result, XMLParseError::Ok);
    
    // 分块发送数据
    std::string xml_part1 = "<?xml version=\"1.0\"?><root><element>";
    std::string xml_part2 = "content</element><another/></root>";
    
    result = reader_->feedData(xml_part1.c_str(), xml_part1.size());
    EXPECT_EQ(result, XMLParseError::Ok);
    
    result = reader_->feedData(xml_part2.c_str(), xml_part2.size());
    EXPECT_EQ(result, XMLParseError::Ok);
    
    result = reader_->endParsing();
    EXPECT_EQ(result, XMLParseError::Ok);
    
    // 验证解析结果
    EXPECT_GE(elements.size(), 3);  // root, element, another
}

// 测试5: DOM解析
TEST_F(XMLStreamReaderTest, DOMParsing) {
    LOG_INFO("Testing DOM-style XML parsing");

    auto root = reader_->parseToDOM(simple_xml_);
    ASSERT_NE(root, nullptr);
    
    EXPECT_EQ(root->name, "root");
    EXPECT_GT(root->children.size(), 0);
    
    // 查找特定元素
    auto element = root->findChild("element");
    ASSERT_NE(element, nullptr);
    EXPECT_EQ(element->getAttribute("attr"), "value");
    EXPECT_EQ(element->getTextContent(), "Text content");
    
    // 查找父元素
    auto parent = root->findChild("parent");
    ASSERT_NE(parent, nullptr);
    
    // 查找子元素
    auto children = parent->findChildren("child");
    EXPECT_EQ(children.size(), 2);
}

// 测试6: 错误处理
TEST_F(XMLStreamReaderTest, ErrorHandling) {
    LOG_INFO("Testing XML parsing error handling");

    std::string invalid_xml = "<?xml version=\"1.0\"?><root><unclosed>";
    
    bool error_callback_called = false;
    reader_->setErrorCallback([&](XMLParseError error, const std::string& message, int line, int column) {
        error_callback_called = true;
        LOG_DEBUG("Parse error: {} at line {}, column {}: {}", static_cast<int>(error), line, column, message);
    });
    
    XMLParseError result = reader_->parseFromString(invalid_xml);
    EXPECT_NE(result, XMLParseError::Ok);
    EXPECT_TRUE(error_callback_called);
}

// 测试7: 大文件处理（模拟）
TEST_F(XMLStreamReaderTest, LargeFileParsing) {
    LOG_INFO("Testing large XML file parsing simulation");

    // 生成一个较大的XML文档
    std::ostringstream large_xml;
    large_xml << "<?xml version=\"1.0\"?><root>";
    
    for (int i = 0; i < 1000; ++i) {
        large_xml << "<item id=\"" << i << "\">Content " << i << "</item>";
    }
    
    large_xml << "</root>";
    
    std::string xml_content = large_xml.str();
    
    int element_count = 0;
    reader_->setStartElementCallback([&](const std::string& name, const std::vector<XMLAttribute>& attributes, int depth) {
        element_count++;
    });
    
    XMLParseError result = reader_->parseFromString(xml_content);
    EXPECT_EQ(result, XMLParseError::Ok);
    EXPECT_EQ(element_count, 1001);  // root + 1000 items
    
    LOG_DEBUG("Successfully parsed {} elements from large XML", element_count);
}

// 测试8: 编码处理
TEST_F(XMLStreamReaderTest, EncodingHandling) {
    LOG_INFO("Testing XML encoding handling");

    std::string utf8_xml = R"(<?xml version="1.0" encoding="UTF-8"?>
<root>
    <text>Hello 世界</text>
    <emoji>🚀</emoji>
</root>)";
    
    std::vector<std::string> texts;
    reader_->setTextCallback([&](const std::string& text, int depth) {
        if (!text.empty()) {
            texts.push_back(text);
        }
    });
    
    reader_->setEncoding("UTF-8");
    XMLParseError result = reader_->parseFromString(utf8_xml);
    EXPECT_EQ(result, XMLParseError::Ok);
    
    // 验证UTF-8文本被正确解析
    bool found_chinese = false;
    bool found_emoji = false;
    for (const auto& text : texts) {
        if (text.find("世界") != std::string::npos) {
            found_chinese = true;
        }
        if (text.find("🚀") != std::string::npos) {
            found_emoji = true;
        }
    }
    EXPECT_TRUE(found_chinese);
    EXPECT_TRUE(found_emoji);
}

// 测试9: 命名空间处理
TEST_F(XMLStreamReaderTest, NamespaceHandling) {
    LOG_INFO("Testing XML namespace handling");

    std::string ns_xml = R"(<?xml version="1.0"?>
<root xmlns:xl="http://www.w3.org/1999/xlink">
    <xl:element xl:attr="value">Content</xl:element>
</root>)";
    
    std::vector<std::string> elements;
    reader_->setStartElementCallback([&](const std::string& name, const std::vector<XMLAttribute>& attributes, int depth) {
        elements.push_back(name);
        LOG_DEBUG("Element with namespace: {}", name);
    });
    
    reader_->setNamespaceAware(true);
    XMLParseError result = reader_->parseFromString(ns_xml);
    EXPECT_EQ(result, XMLParseError::Ok);
    
    // 验证命名空间元素被正确解析
    bool found_ns_element = false;
    for (const auto& element : elements) {
        if (element.find("xl") != std::string::npos) {
            found_ns_element = true;
            break;
        }
    }
    EXPECT_TRUE(found_ns_element);
}

// 测试10: 性能测试
TEST_F(XMLStreamReaderTest, PerformanceTest) {
    LOG_INFO("Testing XML parsing performance");

    // 生成一个中等大小的XML文档
    std::ostringstream perf_xml;
    perf_xml << "<?xml version=\"1.0\"?><workbook>";
    
    for (int sheet = 1; sheet <= 10; ++sheet) {
        perf_xml << "<worksheet name=\"Sheet" << sheet << "\">";
        for (int row = 1; row <= 100; ++row) {
            perf_xml << "<row r=\"" << row << "\">";
            for (int col = 1; col <= 10; ++col) {
                perf_xml << "<c r=\"" << static_cast<char>('A' + col - 1) << row 
                        << "\" t=\"inlineStr\"><is><t>Cell " << row << "," << col << "</t></is></c>";
            }
            perf_xml << "</row>";
        }
        perf_xml << "</worksheet>";
    }
    
    perf_xml << "</workbook>";
    
    std::string xml_content = perf_xml.str();
    LOG_DEBUG("Generated XML content size: {} bytes", xml_content.size());
    
    int element_count = 0;
    reader_->setStartElementCallback([&](const std::string& name, const std::vector<XMLAttribute>& attributes, int depth) {
        element_count++;
    });
    
    auto start_time = std::chrono::high_resolution_clock::now();
    XMLParseError result = reader_->parseFromString(xml_content);
    auto end_time = std::chrono::high_resolution_clock::now();
    
    EXPECT_EQ(result, XMLParseError::Ok);
    
    auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time);
    LOG_INFO("Parsed {} elements in {} ms", element_count, duration.count());
    
    // 性能应该是合理的（这个测试主要是为了确保没有明显的性能问题）
    EXPECT_LT(duration.count(), 5000);  // 应该在5秒内完成
}

}} // namespace fastexcel::xml
