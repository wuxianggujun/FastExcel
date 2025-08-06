/**
 * @file test_optimization.cpp
 * @brief 优化功能单元测试
 * 
 * 测试TimeUtils工具类和统一接口设计的功能
 */

#include <gtest/gtest.h>
#include "fastexcel/utils/TimeUtils.hpp"
#include "fastexcel/core/IFileWriter.hpp"
#include "fastexcel/core/BatchFileWriter.hpp"
#include "fastexcel/core/StreamingFileWriter.hpp"
#include <chrono>
#include <thread>

using namespace fastexcel::utils;
using namespace fastexcel::core;

// ========== TimeUtils 测试 ==========

class TimeUtilsTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 测试前准备
    }
    
    void TearDown() override {
        // 测试后清理
    }
};

TEST_F(TimeUtilsTest, GetCurrentTime) {
    auto current_time = TimeUtils::getCurrentTime();
    
    // 验证时间是合理的（年份应该在合理范围内）
    EXPECT_GE(current_time.tm_year + 1900, 2020);
    EXPECT_LE(current_time.tm_year + 1900, 2030);
    
    // 验证月份范围
    EXPECT_GE(current_time.tm_mon, 0);
    EXPECT_LE(current_time.tm_mon, 11);
    
    // 验证日期范围
    EXPECT_GE(current_time.tm_mday, 1);
    EXPECT_LE(current_time.tm_mday, 31);
}

TEST_F(TimeUtilsTest, GetCurrentUTCTime) {
    auto utc_time = TimeUtils::getCurrentUTCTime();
    
    // UTC时间应该也在合理范围内
    EXPECT_GE(utc_time.tm_year + 1900, 2020);
    EXPECT_LE(utc_time.tm_year + 1900, 2030);
}

TEST_F(TimeUtilsTest, FormatTimeISO8601) {
    // 创建一个特定的时间
    auto test_time = TimeUtils::createTime(2024, 8, 6, 9, 15, 30);
    
    std::string iso_string = TimeUtils::formatTimeISO8601(test_time);
    
    // 验证ISO 8601格式
    EXPECT_EQ(iso_string, "2024-08-06T09:15:30Z");
}

TEST_F(TimeUtilsTest, FormatTimeCustom) {
    auto test_time = TimeUtils::createTime(2024, 8, 6, 9, 15, 30);
    
    std::string custom_format = TimeUtils::formatTime(test_time, "%Y年%m月%d日 %H:%M:%S");
    
    // 验证自定义格式
    EXPECT_EQ(custom_format, "2024年08月06日 09:15:30");
}

TEST_F(TimeUtilsTest, CreateTime) {
    auto test_time = TimeUtils::createTime(2024, 8, 6, 9, 15, 30);
    
    EXPECT_EQ(test_time.tm_year + 1900, 2024);
    EXPECT_EQ(test_time.tm_mon + 1, 8);
    EXPECT_EQ(test_time.tm_mday, 6);
    EXPECT_EQ(test_time.tm_hour, 9);
    EXPECT_EQ(test_time.tm_min, 15);
    EXPECT_EQ(test_time.tm_sec, 30);
}

TEST_F(TimeUtilsTest, ToExcelSerialNumber) {
    // 测试Excel序列号转换
    auto test_time = TimeUtils::createTime(1900, 1, 1); // Excel起始日期
    double serial = TimeUtils::toExcelSerialNumber(test_time);
    
    // Excel序列号应该是1（1900年1月1日）
    EXPECT_DOUBLE_EQ(serial, 1.0);
    
    // 测试另一个日期
    auto test_time2 = TimeUtils::createTime(1900, 1, 2);
    double serial2 = TimeUtils::toExcelSerialNumber(test_time2);
    EXPECT_DOUBLE_EQ(serial2, 2.0);
}

TEST_F(TimeUtilsTest, DaysBetween) {
    auto start_time = TimeUtils::createTime(2024, 8, 1);
    auto end_time = TimeUtils::createTime(2024, 8, 6);
    
    int days = TimeUtils::daysBetween(start_time, end_time);
    EXPECT_EQ(days, 5);
    
    // 测试反向
    int days_reverse = TimeUtils::daysBetween(end_time, start_time);
    EXPECT_EQ(days_reverse, -5);
}

TEST_F(TimeUtilsTest, PerformanceTimer) {
    {
        TimeUtils::PerformanceTimer timer("测试计时器");
        
        // 模拟一些工作
        std::this_thread::sleep_for(std::chrono::milliseconds(100));
        
        // 检查已过时间
        int64_t elapsed = timer.elapsedMs();
        EXPECT_GE(elapsed, 90);  // 至少90ms
        EXPECT_LE(elapsed, 150); // 最多150ms（考虑系统误差）
    }
    // 析构时会自动输出时间（在实际使用中）
}

TEST_F(TimeUtilsTest, GetTimestampMs) {
    int64_t timestamp1 = TimeUtils::getTimestampMs();
    std::this_thread::sleep_for(std::chrono::milliseconds(10));
    int64_t timestamp2 = TimeUtils::getTimestampMs();
    
    // 第二个时间戳应该大于第一个
    EXPECT_GT(timestamp2, timestamp1);
    
    // 差值应该大约是10ms
    int64_t diff = timestamp2 - timestamp1;
    EXPECT_GE(diff, 8);  // 至少8ms
    EXPECT_LE(diff, 20); // 最多20ms
}

// ========== 模拟FileManager用于测试 ==========

class MockFileManager {
public:
    std::vector<std::pair<std::string, std::string>> written_files;
    std::string current_streaming_file;
    std::string current_streaming_content;
    bool streaming_open = false;
    
    bool writeFile(const std::string& path, const std::string& content) {
        written_files.emplace_back(path, content);
        return true;
    }
    
    bool writeFiles(std::vector<std::pair<std::string, std::string>>&& files) {
        for (auto& file : files) {
            written_files.emplace_back(std::move(file));
        }
        return true;
    }
    
    bool openStreamingFile(const std::string& path) {
        if (streaming_open) return false;
        current_streaming_file = path;
        current_streaming_content.clear();
        streaming_open = true;
        return true;
    }
    
    bool writeStreamingChunk(const char* data, size_t size) {
        if (!streaming_open) return false;
        current_streaming_content.append(data, size);
        return true;
    }
    
    bool closeStreamingFile() {
        if (!streaming_open) return false;
        written_files.emplace_back(current_streaming_file, current_streaming_content);
        streaming_open = false;
        current_streaming_file.clear();
        current_streaming_content.clear();
        return true;
    }
};

// ========== IFileWriter 测试 ==========

class FileWriterTest : public ::testing::Test {
protected:
    std::unique_ptr<MockFileManager> mock_file_manager;
    
    void SetUp() override {
        mock_file_manager = std::make_unique<MockFileManager>();
    }
    
    void TearDown() override {
        mock_file_manager.reset();
    }
};

TEST_F(FileWriterTest, BatchFileWriterBasic) {
    BatchFileWriter writer(reinterpret_cast<archive::FileManager*>(mock_file_manager.get()));
    
    // 测试基本写入
    EXPECT_TRUE(writer.writeFile("test1.xml", "<xml>content1</xml>"));
    EXPECT_TRUE(writer.writeFile("test2.xml", "<xml>content2</xml>"));
    
    // 验证文件数量
    EXPECT_EQ(writer.getFileCount(), 2);
    
    // 验证类型名称
    EXPECT_EQ(writer.getTypeName(), "BatchFileWriter");
    
    // 验证统计信息
    auto stats = writer.getStats();
    EXPECT_EQ(stats.batch_files, 2);
    EXPECT_GT(stats.total_bytes, 0);
}

TEST_F(FileWriterTest, BatchFileWriterStreaming) {
    BatchFileWriter writer(reinterpret_cast<archive::FileManager*>(mock_file_manager.get()));
    
    // 测试流式写入（在批量模式中会收集到内存）
    EXPECT_TRUE(writer.openStreamingFile("streaming_test.xml"));
    EXPECT_TRUE(writer.writeStreamingChunk("<xml>", 5));
    EXPECT_TRUE(writer.writeStreamingChunk("content", 7));
    EXPECT_TRUE(writer.writeStreamingChunk("</xml>", 6));
    EXPECT_TRUE(writer.closeStreamingFile());
    
    // 验证统计信息
    auto stats = writer.getStats();
    EXPECT_EQ(stats.streaming_files, 1);
    EXPECT_EQ(stats.batch_files, 1); // 流式写入在批量模式中也算作批量文件
}

TEST_F(FileWriterTest, StreamingFileWriterBasic) {
    StreamingFileWriter writer(reinterpret_cast<archive::FileManager*>(mock_file_manager.get()));
    
    // 验证类型名称
    EXPECT_EQ(writer.getTypeName(), "StreamingFileWriter");
    
    // 验证初始状态
    EXPECT_FALSE(writer.hasOpenStreamingFile());
    EXPECT_TRUE(writer.getCurrentStreamingPath().empty());
}

// ========== 集成测试 ==========

class OptimizationIntegrationTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 集成测试准备
    }
};

TEST_F(OptimizationIntegrationTest, TimeUtilsWithFileWriter) {
    // 测试时间工具类与文件写入器的集成
    auto current_time = TimeUtils::getCurrentTime();
    std::string time_string = TimeUtils::formatTimeISO8601(current_time);
    
    MockFileManager mock_manager;
    BatchFileWriter writer(reinterpret_cast<archive::FileManager*>(&mock_manager));
    
    // 创建包含时间信息的XML内容
    std::string xml_content = "<?xml version=\"1.0\"?>\n<document created=\"" + time_string + "\"/>";
    
    EXPECT_TRUE(writer.writeFile("document.xml", xml_content));
    
    auto stats = writer.getStats();
    EXPECT_EQ(stats.batch_files, 1);
    EXPECT_GT(stats.total_bytes, time_string.length());
}

// ========== 性能测试 ==========

class PerformanceTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 性能测试准备
    }
};

TEST_F(PerformanceTest, TimeUtilsPerformance) {
    const int iterations = 10000;
    
    TimeUtils::PerformanceTimer timer("TimeUtils性能测试");
    
    for (int i = 0; i < iterations; ++i) {
        auto time = TimeUtils::getCurrentTime();
        auto iso_string = TimeUtils::formatTimeISO8601(time);
        auto excel_serial = TimeUtils::toExcelSerialNumber(time);
        
        // 防止编译器优化掉这些调用
        (void)iso_string;
        (void)excel_serial;
    }
    
    int64_t elapsed = timer.elapsedMs();
    
    // 10000次调用应该在合理时间内完成（比如1秒内）
    EXPECT_LT(elapsed, 1000);
    
    std::cout << "TimeUtils性能测试: " << iterations << " 次调用耗时 " << elapsed << " ms" << std::endl;
}