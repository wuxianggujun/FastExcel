#include <gtest/gtest.h>
#include "fastexcel/xml/XMLStreamWriter.hpp"
#include <sstream>
#include <fstream>
#include <filesystem>
#include <chrono>
#include <thread>
#include <vector>
#include <atomic>

using namespace fastexcel::xml;

class XMLStreamWriterTest : public ::testing::Test {
protected:
    void SetUp() override {
        writer = std::make_unique<XMLStreamWriter>();
        test_filename = "test_output.xml";
    }
    
    void TearDown() override {
        writer.reset();
        // 清理测试文件
        if (std::filesystem::exists(test_filename)) {
            std::filesystem::remove(test_filename);
        }
    }
    
    std::unique_ptr<XMLStreamWriter> writer;
    std::string test_filename;
};

// 测试基本XML文档生成
TEST_F(XMLStreamWriterTest, BasicDocument) {
    writer->startDocument();
    writer->startElement("root");
    writer->writeText("Hello World");
    writer->endElement();
    writer->endDocument();
    
    std::string xml = writer->toString();
    EXPECT_FALSE(xml.empty());
    EXPECT_NE(xml.find("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"), std::string::npos);
    EXPECT_NE(xml.find("<root>"), std::string::npos);
    EXPECT_NE(xml.find("Hello World"), std::string::npos);
    EXPECT_NE(xml.find("</root>"), std::string::npos);
}

// 测试元素嵌套
TEST_F(XMLStreamWriterTest, NestedElements) {
    writer->startDocument();
    writer->startElement("parent");
    writer->startElement("child1");
    writer->writeText("Child 1 Content");
    writer->endElement();
    writer->startElement("child2");
    writer->writeText("Child 2 Content");
    writer->endElement();
    writer->endElement();
    writer->endDocument();
    
    std::string xml = writer->toString();
    EXPECT_NE(xml.find("<parent>"), std::string::npos);
    EXPECT_NE(xml.find("<child1>"), std::string::npos);
    EXPECT_NE(xml.find("Child 1 Content"), std::string::npos);
    EXPECT_NE(xml.find("</child1>"), std::string::npos);
    EXPECT_NE(xml.find("<child2>"), std::string::npos);
    EXPECT_NE(xml.find("Child 2 Content"), std::string::npos);
    EXPECT_NE(xml.find("</child2>"), std::string::npos);
    EXPECT_NE(xml.find("</parent>"), std::string::npos);
}

// 测试属性写入
TEST_F(XMLStreamWriterTest, Attributes) {
    writer->startDocument();
    writer->startElement("element");
    writer->writeAttribute("attr1", "value1");
    writer->writeAttribute("attr2", "value2");
    writer->writeText("Content");
    writer->endElement();
    writer->endDocument();
    
    std::string xml = writer->toString();
    EXPECT_NE(xml.find("attr1=\"value1\""), std::string::npos);
    EXPECT_NE(xml.find("attr2=\"value2\""), std::string::npos);
    EXPECT_NE(xml.find("Content"), std::string::npos);
}

// 测试数字属性
TEST_F(XMLStreamWriterTest, NumericAttributes) {
    writer->startDocument();
    writer->startElement("element");
    writer->writeAttribute("intAttr", 42);
    writer->writeAttribute("doubleAttr", 3.14159);
    writer->endElement();
    writer->endDocument();
    
    std::string xml = writer->toString();
    EXPECT_NE(xml.find("intAttr=\"42\""), std::string::npos);
    EXPECT_NE(xml.find("doubleAttr=\"3.14159\""), std::string::npos);
}

// 测试空元素
TEST_F(XMLStreamWriterTest, EmptyElements) {
    writer->startDocument();
    writer->writeEmptyElement("empty1");
    writer->startElement("parent");
    writer->writeEmptyElement("empty2");
    writer->endElement();
    writer->endDocument();
    
    std::string xml = writer->toString();
    EXPECT_NE(xml.find("<empty1/>"), std::string::npos);
    EXPECT_NE(xml.find("<empty2/>"), std::string::npos);
}

// 测试自闭合元素（带属性）
TEST_F(XMLStreamWriterTest, SelfClosingWithAttributes) {
    writer->startDocument();
    writer->startElement("element");
    writer->writeAttribute("attr", "value");
    writer->endElement(); // 应该生成自闭合标签
    writer->endDocument();
    
    std::string xml = writer->toString();
    EXPECT_NE(xml.find("<element attr=\"value\"/>"), std::string::npos);
}

// 测试字符转义
TEST_F(XMLStreamWriterTest, CharacterEscaping) {
    writer->startDocument();
    writer->startElement("test");
    writer->writeAttribute("attr", "value with & < > \" ' characters");
    writer->writeText("Text with & < > characters");
    writer->endElement();
    writer->endDocument();
    
    std::string xml = writer->toString();
    // 属性中的转义
    EXPECT_NE(xml.find("&amp;"), std::string::npos);
    EXPECT_NE(xml.find("&lt;"), std::string::npos);
    EXPECT_NE(xml.find("&gt;"), std::string::npos);
    EXPECT_NE(xml.find("&quot;"), std::string::npos);
    EXPECT_NE(xml.find("&apos;"), std::string::npos);
}

// 测试换行符转义
TEST_F(XMLStreamWriterTest, NewlineEscaping) {
    writer->startDocument();
    writer->startElement("test");
    writer->writeAttribute("attr", "line1\nline2");
    writer->writeText("line1\nline2");
    writer->endElement();
    writer->endDocument();
    
    std::string xml = writer->toString();
    // 属性中的换行符应该被转义
    EXPECT_NE(xml.find("&#xA;"), std::string::npos);
}

// 测试原始数据写入
TEST_F(XMLStreamWriterTest, RawData) {
    writer->startDocument();
    writer->startElement("root");
    writer->writeRaw("<custom>Raw XML Content</custom>");
    writer->endElement();
    writer->endDocument();
    
    std::string xml = writer->toString();
    EXPECT_NE(xml.find("<custom>Raw XML Content</custom>"), std::string::npos);
}

// 测试原始数据写入（string版本）
TEST_F(XMLStreamWriterTest, RawDataString) {
    std::string raw_content = "<item id=\"1\">Content</item>";
    
    writer->startDocument();
    writer->startElement("root");
    writer->writeRaw(raw_content);
    writer->endElement();
    writer->endDocument();
    
    std::string xml = writer->toString();
    EXPECT_NE(xml.find(raw_content), std::string::npos);
}

// 测试缓冲模式
TEST_F(XMLStreamWriterTest, BufferedMode) {
    writer->setBufferedMode();
    
    writer->startDocument();
    writer->startElement("test");
    writer->writeText("Buffered content");
    writer->endElement();
    writer->endDocument();
    
    std::string xml = writer->toString();
    EXPECT_FALSE(xml.empty());
    EXPECT_NE(xml.find("Buffered content"), std::string::npos);
}

// 测试文件写入模式
TEST_F(XMLStreamWriterTest, FileMode) {
    FILE* file = nullptr;
    errno_t err = fopen_s(&file, test_filename.c_str(), "wb");
    ASSERT_EQ(err, 0);
    ASSERT_NE(file, nullptr);
    
    writer->setDirectFileMode(file, true); // 让writer拥有文件
    
    writer->startDocument();
    writer->startElement("fileTest");
    writer->writeText("File content");
    writer->endElement();
    writer->endDocument();
    
    // 重置writer以确保文件被关闭
    writer.reset();
    
    // 验证文件内容
    std::ifstream input(test_filename);
    std::string content((std::istreambuf_iterator<char>(input)),
                       std::istreambuf_iterator<char>());
    
    EXPECT_FALSE(content.empty());
    EXPECT_NE(content.find("File content"), std::string::npos);
}

// 测试清空操作
TEST_F(XMLStreamWriterTest, Clear) {
    writer->startDocument();
    writer->startElement("test");
    writer->writeText("Some content");
    writer->endElement();
    writer->endDocument();
    
    std::string xml1 = writer->toString();
    EXPECT_FALSE(xml1.empty());
    
    writer->clear();
    std::string xml2 = writer->toString();
    EXPECT_TRUE(xml2.empty());
    
    // 清空后应该能重新使用
    writer->startDocument();
    writer->startElement("new");
    writer->writeText("New content");
    writer->endElement();
    writer->endDocument();
    
    std::string xml3 = writer->toString();
    EXPECT_FALSE(xml3.empty());
    EXPECT_NE(xml3.find("New content"), std::string::npos);
    EXPECT_EQ(xml3.find("Some content"), std::string::npos);
}

// 测试属性批处理
TEST_F(XMLStreamWriterTest, AttributeBatch) {
    writer->startDocument();
    writer->startElement("element");
    
    writer->startAttributeBatch();
    writer->writeAttribute("attr1", "value1");
    writer->writeAttribute("attr2", "value2");
    writer->writeAttribute("attr3", "value3");
    writer->endAttributeBatch();
    
    writer->writeText("Content");
    writer->endElement();
    writer->endDocument();
    
    std::string xml = writer->toString();
    EXPECT_NE(xml.find("attr1=\"value1\""), std::string::npos);
    EXPECT_NE(xml.find("attr2=\"value2\""), std::string::npos);
    EXPECT_NE(xml.find("attr3=\"value3\""), std::string::npos);
}

// 测试大量数据
TEST_F(XMLStreamWriterTest, LargeData) {
    writer->startDocument();
    writer->startElement("root");
    
    // 写入大量元素
    for (int i = 0; i < 1000; ++i) {
        writer->startElement("item");
        writer->writeAttribute("id", i);
        writer->writeText(("Item " + std::to_string(i)).c_str());
        writer->endElement();
    }
    
    writer->endElement();
    writer->endDocument();
    
    std::string xml = writer->toString();
    EXPECT_FALSE(xml.empty());
    EXPECT_GT(xml.length(), 10000); // 应该是一个相当大的XML
    
    // 验证第一个和最后一个元素
    EXPECT_NE(xml.find("<item id=\"0\">Item 0</item>"), std::string::npos);
    EXPECT_NE(xml.find("<item id=\"999\">Item 999</item>"), std::string::npos);
}

// 测试错误处理
TEST_F(XMLStreamWriterTest, ErrorHandling) {
    // 在没有开始元素的情况下写入属性应该被忽略或处理
    writer->writeAttribute("attr", "value");
    
    // 在没有开始文档的情况下结束文档
    writer->endDocument();
    
    // 这些操作不应该崩溃
    std::string xml = writer->toString();
    // XML可能为空或包含部分内容，但不应该崩溃
}

// 测试元素栈管理
TEST_F(XMLStreamWriterTest, ElementStack) {
    writer->startDocument();
    writer->startElement("level1");
    writer->startElement("level2");
    writer->startElement("level3");
    writer->writeText("Deep content");
    writer->endElement(); // level3
    writer->endElement(); // level2
    writer->endElement(); // level1
    writer->endDocument();
    
    std::string xml = writer->toString();
    EXPECT_NE(xml.find("<level1>"), std::string::npos);
    EXPECT_NE(xml.find("<level2>"), std::string::npos);
    EXPECT_NE(xml.find("<level3>"), std::string::npos);
    EXPECT_NE(xml.find("Deep content"), std::string::npos);
    EXPECT_NE(xml.find("</level3>"), std::string::npos);
    EXPECT_NE(xml.find("</level2>"), std::string::npos);
    EXPECT_NE(xml.find("</level1>"), std::string::npos);
}

// 测试空文本处理
TEST_F(XMLStreamWriterTest, EmptyText) {
    writer->startDocument();
    writer->startElement("test");
    writer->writeText("");
    writer->endElement();
    writer->endDocument();
    
    std::string xml = writer->toString();
    EXPECT_NE(xml.find("<test></test>"), std::string::npos);
}

// 测试空属性值
TEST_F(XMLStreamWriterTest, EmptyAttributeValue) {
    writer->startDocument();
    writer->startElement("test");
    writer->writeAttribute("empty", "");
    writer->writeAttribute("normal", "value");
    writer->endElement();
    writer->endDocument();
    
    std::string xml = writer->toString();
    EXPECT_NE(xml.find("empty=\"\""), std::string::npos);
    EXPECT_NE(xml.find("normal=\"value\""), std::string::npos);
}

// 测试特殊字符处理
TEST_F(XMLStreamWriterTest, SpecialCharacters) {
    writer->startDocument();
    writer->startElement("test");
    writer->writeText("Unicode: 中文 العربية русский");
    writer->endElement();
    writer->endDocument();
    
    std::string xml = writer->toString();
    EXPECT_NE(xml.find("中文"), std::string::npos);
    EXPECT_NE(xml.find("العربية"), std::string::npos);
    EXPECT_NE(xml.find("русский"), std::string::npos);
}

// 测试性能（基本测试）
TEST_F(XMLStreamWriterTest, Performance) {
    auto start = std::chrono::high_resolution_clock::now();
    
    writer->startDocument();
    writer->startElement("root");
    
    // 生成大量XML内容
    for (int i = 0; i < 10000; ++i) {
        writer->startElement("item");
        writer->writeAttribute("id", i);
        writer->writeAttribute("name", ("Item " + std::to_string(i)).c_str());
        writer->writeText(("Content " + std::to_string(i)).c_str());
        writer->endElement();
    }
    
    writer->endElement();
    writer->endDocument();
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);
    
    // 性能测试：10000个元素应该在合理时间内完成
    EXPECT_LT(duration.count(), 1000); // 少于1秒
    
    std::string xml = writer->toString();
    EXPECT_FALSE(xml.empty());
}

// 测试内存使用
TEST_F(XMLStreamWriterTest, MemoryUsage) {
    // 这是一个基本的内存使用测试
    writer->startDocument();
    writer->startElement("root");
    
    // 生成大量内容然后清空，测试内存是否正确释放
    for (int batch = 0; batch < 10; ++batch) {
        for (int i = 0; i < 1000; ++i) {
            writer->startElement("temp");
            writer->writeText(("Temporary content " + std::to_string(i)).c_str());
            writer->endElement();
        }
        
        // 获取当前XML大小
        std::string xml = writer->toString();
        size_t size = xml.length();
        
        // 清空并重新开始
        writer->clear();
        writer->startDocument();
        writer->startElement("root");
        
        // 验证清空后内存使用减少
        std::string empty_xml = writer->toString();
        EXPECT_LT(empty_xml.length(), size);
    }
    
    writer->endElement();
    writer->endDocument();
}

// 测试并发安全（基本测试）
TEST_F(XMLStreamWriterTest, ThreadSafety) {
    // 注意：XMLStreamWriter可能不是线程安全的，这个测试主要是验证不会崩溃
    std::vector<std::thread> threads;
    std::atomic<int> success_count(0);
    
    for (int i = 0; i < 3; ++i) {
        threads.emplace_back([this, i, &success_count]() {
            try {
                auto local_writer = std::make_unique<XMLStreamWriter>();
                local_writer->startDocument();
                local_writer->startElement(("thread" + std::to_string(i)).c_str());
                local_writer->writeText(("Thread " + std::to_string(i) + " content").c_str());
                local_writer->endElement();
                local_writer->endDocument();
                
                std::string xml = local_writer->toString();
                if (!xml.empty()) {
                    success_count++;
                }
            } catch (...) {
                // 捕获任何异常
            }
        });
    }
    
    for (auto& thread : threads) {
        thread.join();
    }
    
    // 所有线程都应该成功（因为使用了独立的writer实例）
    EXPECT_EQ(success_count.load(), 3);
}