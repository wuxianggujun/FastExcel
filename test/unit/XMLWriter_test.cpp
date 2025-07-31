#include "fastexcel/xml/XMLStreamWriter.hpp"
#include <gtest/gtest.h>
#include <string>
#include <chrono>
#include <cstdio>
#include <filesystem>

namespace fastexcel {
namespace xml {

class XMLStreamWriterTest : public ::testing::Test {
protected:
    void SetUp() override {
        writer = std::make_unique<XMLStreamWriter>();
        test_dir_ = "test_output";
        std::filesystem::create_directories(test_dir_);
    }

    void TearDown() override {
        writer.reset();
        // 清理测试文件
        std::filesystem::remove_all(test_dir_);
    }

    std::unique_ptr<XMLStreamWriter> writer;
    std::string test_dir_;
};

TEST_F(XMLStreamWriterTest, BasicDocumentCreation) {
    writer->startDocument();
    writer->startElement("root");
    writer->writeAttribute("version", "1.0");
    writer->writeText("Hello World");
    writer->endElement();
    writer->endDocument();
    
    std::string result = writer->toString();
    std::string expected = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<root version=\"1.0\">Hello World</root>";
    
    EXPECT_EQ(result, expected);
}

TEST_F(XMLStreamWriterTest, EmptyElement) {
    writer->startDocument();
    writer->writeEmptyElement("empty");
    writer->endDocument();
    
    std::string result = writer->toString();
    std::string expected = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<empty/>";
    
    EXPECT_EQ(result, expected);
}

TEST_F(XMLStreamWriterTest, NestedElements) {
    writer->startDocument();
    writer->startElement("root");
    writer->startElement("child");
    writer->writeAttribute("attr", "value");
    writer->writeText("content");
    writer->endElement();
    writer->endElement();
    writer->endDocument();
    
    std::string result = writer->toString();
    std::string expected = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<root><child attr=\"value\">content</child></root>";
    
    EXPECT_EQ(result, expected);
}

TEST_F(XMLStreamWriterTest, TextEscaping) {
    writer->startDocument();
    writer->startElement("test");
    writer->writeText("Special chars: < > & \" '");
    writer->endElement();
    writer->endDocument();
    
    std::string result = writer->toString();
    std::string expected = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<test>Special chars: < > & \" '</test>";
    
    EXPECT_EQ(result, expected);
}

TEST_F(XMLStreamWriterTest, AttributeEscaping) {
    writer->startDocument();
    writer->startElement("test");
    writer->writeAttribute("attr", "Special: < > & \" '");
    writer->endElement();
    writer->endDocument();
    
    std::string result = writer->toString();
    std::string expected = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<test attr=\"Special: < > & " '\"/>";
    
    EXPECT_EQ(result, expected);
}

TEST_F(XMLStreamWriterTest, NumericAttributes) {
    writer->startDocument();
    writer->startElement("test");
    writer->writeAttribute("int", 42);
    writer->writeAttribute("double", 3.14159);
    writer->endElement();
    writer->endDocument();
    
    std::string result = writer->toString();
    EXPECT_TRUE(result.find("int=\"42\"") != std::string::npos);
    EXPECT_TRUE(result.find("double=\"3.14159\"") != std::string::npos);
}

TEST_F(XMLStreamWriterTest, NewlineEscapingInAttributes) {
    writer->startDocument();
    writer->startElement("test");
    writer->writeAttribute("attr", "Line1\nLine2");
    writer->endElement();
    writer->endDocument();
    
    std::string result = writer->toString();
    EXPECT_TRUE(result.find("attr=\"Line1&#xA;Line2\"") != std::string::npos);
}

TEST_F(XMLStreamWriterTest, PerformanceTest) {
    const int iterations = 10000;
    
    auto start = std::chrono::high_resolution_clock::now();
    
    for (int i = 0; i < iterations; ++i) {
        writer->clear();
        writer->startDocument();
        writer->startElement("root");
        writer->writeAttribute("id", std::to_string(i));
        writer->startElement("item");
        writer->writeAttribute("name", "item_" + std::to_string(i));
        writer->writeText("Content for item " + std::to_string(i));
        writer->endElement();
        writer->endElement();
        writer->endDocument();
    }
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);
    
    LOG_INFO("Performance test: {} iterations in {} ms ({} ms per iteration)", 
             iterations, duration.count(), static_cast<double>(duration.count()) / iterations);
    
    // 性能应该足够快，每个迭代应该在1ms以内
    EXPECT_LT(duration.count(), iterations * 1);
}

TEST_F(XMLStreamWriterTest, LargeDocumentTest) {
    const int elements = 1000;
    
    writer->startDocument();
    writer->startElement("root");
    
    for (int i = 0; i < elements; ++i) {
        writer->startElement("item");
        writer->writeAttribute("id", std::to_string(i));
        writer->writeText("Item " + std::to_string(i));
        writer->endElement();
    }
    
    writer->endElement();
    writer->endDocument();
    
    std::string result = writer->toString();
    
    // 验证文档结构
    EXPECT_TRUE(result.find("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>") == 0);
    EXPECT_TRUE(result.find("<root>") != std::string::npos);
    EXPECT_TRUE(result.find("</root>") != std::string::npos);
    
    // 验证元素数量
    size_t item_count = 0;
    size_t pos = 0;
    while ((pos = result.find("<item", pos)) != std::string::npos) {
        item_count++;
        pos += 5;
    }
    
    EXPECT_EQ(item_count, elements);
}

TEST_F(XMLStreamWriterTest, BufferManagement) {
    // 测试缓冲区自动增长
    std::string large_text(10000, 'x');
    
    writer->startDocument();
    writer->startElement("root");
    writer->writeText(large_text);
    writer->endElement();
    writer->endDocument();
    
    std::string result = writer->toString();
    EXPECT_TRUE(result.find(large_text) != std::string::npos);
}

TEST_F(XMLStreamWriterTest, ClearAndReuse) {
    // 第一次使用
    writer->startDocument();
    writer->startElement("first");
    writer->writeText("First content");
    writer->endElement();
    writer->endDocument();
    
    std::string first_result = writer->toString();
    
    // 清除并重用
    writer->clear();
    
    // 第二次使用
    writer->startDocument();
    writer->startElement("second");
    writer->writeText("Second content");
    writer->endElement();
    writer->endDocument();
    
    std::string second_result = writer->toString();
    
    // 验证第二次结果不包含第一次的内容
    EXPECT_TRUE(first_result.find("first") != std::string::npos);
    EXPECT_TRUE(second_result.find("second") != std::string::npos);
    EXPECT_TRUE(second_result.find("first") == std::string::npos);
}

TEST_F(XMLStreamWriterTest, FileOutputTest) {
    std::string test_file = test_dir_ + "/test_output.xml";
    
    writer->startDocument();
    writer->startElement("root");
    writer->writeAttribute("version", "1.0");
    writer->writeText("Hello World");
    writer->endElement();
    writer->endDocument();
    
    // 写入文件
    EXPECT_TRUE(writer->writeToFile(test_file));
    
    // 验证文件存在
    EXPECT_TRUE(std::filesystem::exists(test_file));
    
    // 读取文件内容并验证
    std::ifstream file(test_file);
    std::string content((std::istreambuf_iterator<char>(file)), 
                        std::istreambuf_iterator<char>());
    file.close();
    
    std::string expected = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<root version=\"1.0\">Hello World</root>";
    EXPECT_EQ(content, expected);
}

TEST_F(XMLStreamWriterTest, DirectFileModeTest) {
    std::string test_file = test_dir_ + "/test_direct.xml";
    
    // 创建新的XMLStreamWriter用于直接文件模式测试
    auto file_writer = std::make_unique<XMLStreamWriter>();
    
    FILE* file = fopen(test_file.c_str(), "wb");
    ASSERT_NE(file, nullptr);
    
    // 设置直接文件模式
    file_writer->setDirectFileMode(file, true);
    
    // 写入内容
    file_writer->startDocument();
    file_writer->startElement("root");
    file_writer->writeText("Direct file mode content");
    file_writer->endElement();
    file_writer->endDocument();
    
    // 验证文件内容
    std::ifstream input_file(test_file);
    std::string content((std::istreambuf_iterator<char>(input_file)), 
                        std::istreambuf_iterator<char>());
    input_file.close();
    
    std::string expected = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<root>Direct file mode content</root>";
    EXPECT_EQ(content, expected);
}

TEST_F(XMLStreamWriterTest, EscapingOptimizationTest) {
    // 测试不需要转义的文本是否保持原样
    std::string simple_text = "Hello World 12345";
    
    writer->startDocument();
    writer->startElement("test");
    writer->writeText(simple_text);
    writer->endElement();
    writer->endDocument();
    
    std::string result = writer->toString();
    EXPECT_TRUE(result.find(simple_text) != std::string::npos);
    
    // 测试需要转义的文本
    std::string complex_text = "Hello <world> & \"everyone\"";
    
    writer->clear();
    writer->startDocument();
    writer->startElement("test");
    writer->writeText(complex_text);
    writer->endElement();
    writer->endDocument();
    
    result = writer->toString();
    EXPECT_TRUE(result.find("<") != std::string::npos);
    EXPECT_TRUE(result.find(">") != std::string::npos);
    EXPECT_TRUE(result.find("&") != std::string::npos);
    EXPECT_TRUE(result.find(""") != std::string::npos);
}

TEST_F(XMLStreamWriterTest, MemoryEfficiencyTest) {
    // 测试内存重用
    size_t initial_buffer_capacity = 8192; // 固定缓冲区大小
    
    // 写入大量数据
    writer->startDocument();
    writer->startElement("root");
    for (int i = 0; i < 100; ++i) {
        writer->startElement("item");
        writer->writeText("This is a relatively long piece of text to test memory allocation and deallocation patterns");
        writer->endElement();
    }
    writer->endElement();
    writer->endDocument();
    
    // 清除并重用
    writer->clear();
    
    // 写入少量数据
    writer->startDocument();
    writer->startElement("small");
    writer->writeText("Small content");
    writer->endElement();
    writer->endDocument();
    
    // 缓冲区大小应该保持不变
    EXPECT_EQ(initial_buffer_capacity, 8192);
}

TEST_F(XMLStreamWriterTest, PerformanceComparisonTest) {
    // 测试缓冲模式 vs 直接文件模式的性能差异
    std::string test_file_buffered = test_dir_ + "/perf_buffered.xml";
    std::string test_file_direct = test_dir_ + "/perf_direct.xml";
    
    const int elements = 5000;
    
    // 测试缓冲模式
    auto buffered_writer = std::make_unique<XMLStreamWriter>();
    auto start = std::chrono::high_resolution_clock::now();
    
    buffered_writer->startDocument();
    buffered_writer->startElement("root");
    for (int i = 0; i < elements; ++i) {
        buffered_writer->startElement("item");
        buffered_writer->writeAttribute("id", std::to_string(i));
        buffered_writer->writeText("Item " + std::to_string(i));
        buffered_writer->endElement();
    }
    buffered_writer->endElement();
    buffered_writer->endDocument();
    
    EXPECT_TRUE(buffered_writer->writeToFile(test_file_buffered));
    auto end = std::chrono::high_resolution_clock::now();
    auto buffered_duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);
    
    // 测试直接文件模式
    FILE* file = fopen(test_file_direct.c_str(), "wb");
    ASSERT_NE(file, nullptr);
    
    auto direct_writer = std::make_unique<XMLStreamWriter>();
    direct_writer->setDirectFileMode(file, true);
    
    start = std::chrono::high_resolution_clock::now();
    direct_writer->startDocument();
    direct_writer->startElement("root");
    for (int i = 0; i < elements; ++i) {
        direct_writer->startElement("item");
        direct_writer->writeAttribute("id", std::to_string(i));
        direct_writer->writeText("Item " + std::to_string(i));
        direct_writer->endElement();
    }
    direct_writer->endElement();
    direct_writer->endDocument();
    end = std::chrono::high_resolution_clock::now();
    auto direct_duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);
    
    LOG_INFO("Buffered mode: {} ms, Direct mode: {} ms", 
             buffered_duration.count(), direct_duration.count());
    
    // 直接文件模式应该更快（或者至少不慢很多）
    EXPECT_LE(direct_duration.count(), buffered_duration.count() + 100);
}

}} // namespace fastexcel::xml