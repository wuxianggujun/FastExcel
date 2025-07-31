#include "fastexcel/archive/ZipArchive.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <gtest/gtest.h>
#include <filesystem>
#include <fstream>
#include <vector>
#include <thread>
#include <chrono>

namespace fastexcel {
namespace archive {

class ZipArchiveTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 初始化日志系统
        fastexcel::Logger::getInstance().initialize("logs/ZipArchive_test.log", 
                                                   fastexcel::Logger::Level::DEBUG, 
                                                   false);
        
        // 创建测试目录
        test_dir_ = "test_zip_archive";
        std::filesystem::create_directories(test_dir_);
        
        // 创建测试文件路径
        test_zip_path_ = test_dir_ + "/test.zip";
        
        // 确保测试开始时文件不存在
        std::filesystem::remove(test_zip_path_);
    }

    void TearDown() override {
        // 关闭 ZIP 文件
        if (zip_archive_) {
            zip_archive_->close();
        }
        
        // 清理测试文件
        std::filesystem::remove_all(test_dir_);
        
        // 关闭日志系统
        fastexcel::Logger::getInstance().shutdown();
    }

    // 辅助函数：创建测试数据
    std::string createTestString(size_t size = 100) {
        std::string result;
        for (size_t i = 0; i < size; ++i) {
            result += static_cast<char>('A' + (i % 26));
        }
        return result;
    }

    // 辅助函数：创建测试二进制数据
    std::vector<uint8_t> createTestBinaryData(size_t size = 100) {
        std::vector<uint8_t> result(size);
        for (size_t i = 0; i < size; ++i) {
            result[i] = static_cast<uint8_t>(i % 256);
        }
        return result;
    }

    // 测试成员变量
    std::unique_ptr<ZipArchive> zip_archive_;
    std::string test_dir_;
    std::string test_zip_path_;
};

// 测试1: 基本创建和打开
TEST_F(ZipArchiveTest, CreateAndOpen) {
    LOG_INFO("Testing basic ZIP archive creation and opening");

    // 创建 ZIP 文件
    zip_archive_ = std::make_unique<ZipArchive>(test_zip_path_);
    EXPECT_TRUE(zip_archive_->open(true));
    EXPECT_TRUE(zip_archive_->isWritable());
    EXPECT_FALSE(zip_archive_->isReadable());
    
    // 关闭后重新打开
    zip_archive_->close();
    EXPECT_FALSE(zip_archive_->isOpen());
    
    // 重新打开
    EXPECT_TRUE(zip_archive_->open(false));
    EXPECT_FALSE(zip_archive_->isWritable());
    EXPECT_TRUE(zip_archive_->isReadable());
}

// 测试2: 添加和提取字符串文件
TEST_F(ZipArchiveTest, AddAndExtractStringFile) {
    LOG_INFO("Testing adding and extracting string files");

    // 创建 ZIP 文件
    zip_archive_ = std::make_unique<ZipArchive>(test_zip_path_);
    ASSERT_TRUE(zip_archive_->open(true));
    
    // 创建测试数据
    std::string test_content = createTestString(1000);
    std::string internal_path = "test_string.txt";
    
    // 添加文件到 ZIP
    EXPECT_TRUE(zip_archive_->addFile(internal_path, test_content));
    
    // 关闭 ZIP 文件
    zip_archive_->close();
    
    // 确保文件已经完全写入磁盘
    // 短暂延迟以确保文件系统操作完成
    std::this_thread::sleep_for(std::chrono::milliseconds(10));
    
    // 重新打开 ZIP 文件进行读取
    ASSERT_TRUE(zip_archive_->open(false));
    
    // 检查文件是否存在
    EXPECT_TRUE(zip_archive_->fileExists(internal_path));
    
    // 提取文件内容
    std::string extracted_content;
    EXPECT_TRUE(zip_archive_->extractFile(internal_path, extracted_content));
    
    // 验证内容是否一致
    EXPECT_EQ(extracted_content, test_content);
}

// 测试3: 添加和提取二进制文件
TEST_F(ZipArchiveTest, AddAndExtractBinaryFile) {
    LOG_INFO("Testing adding and extracting binary files");

    // 创建 ZIP 文件
    zip_archive_ = std::make_unique<ZipArchive>(test_zip_path_);
    ASSERT_TRUE(zip_archive_->open(true));
    
    // 创建测试数据
    std::vector<uint8_t> test_data = createTestBinaryData(1000);
    std::string internal_path = "test_binary.bin";
    
    // 添加文件到 ZIP
    EXPECT_TRUE(zip_archive_->addFile(internal_path, test_data));
    
    // 关闭 ZIP 文件
    zip_archive_->close();
    
    // 确保文件已经完全写入磁盘
    std::this_thread::sleep_for(std::chrono::milliseconds(10));
    
    // 重新打开 ZIP 文件进行读取
    ASSERT_TRUE(zip_archive_->open(false));
    
    // 检查文件是否存在
    EXPECT_TRUE(zip_archive_->fileExists(internal_path));
    
    // 提取文件内容
    std::vector<uint8_t> extracted_data;
    EXPECT_TRUE(zip_archive_->extractFile(internal_path, extracted_data));
    
    // 验证内容是否一致
    EXPECT_EQ(extracted_data.size(), test_data.size());
    for (size_t i = 0; i < test_data.size(); ++i) {
        EXPECT_EQ(extracted_data[i], test_data[i]);
    }
}

// 测试4: 添加多个文件并列出
TEST_F(ZipArchiveTest, AddMultipleFilesAndList) {
    LOG_INFO("Testing adding multiple files and listing them");

    // 创建 ZIP 文件
    zip_archive_ = std::make_unique<ZipArchive>(test_zip_path_);
    ASSERT_TRUE(zip_archive_->open(true));
    
    // 添加多个文件
    const int file_count = 5;
    for (int i = 0; i < file_count; ++i) {
        std::string path = "file_" + std::to_string(i) + ".txt";
        std::string content = "Content of file " + std::to_string(i);
        EXPECT_TRUE(zip_archive_->addFile(path, content));
    }
    
    // 关闭 ZIP 文件
    zip_archive_->close();
    
    // 确保文件已经完全写入磁盘
    std::this_thread::sleep_for(std::chrono::milliseconds(10));
    
    // 重新打开 ZIP 文件进行读取
    ASSERT_TRUE(zip_archive_->open(false));
    
    // 列出所有文件
    std::vector<std::string> file_list = zip_archive_->listFiles();
    EXPECT_EQ(file_list.size(), file_count);
    
    // 验证所有文件都存在
    for (int i = 0; i < file_count; ++i) {
        std::string path = "file_" + std::to_string(i) + ".txt";
        EXPECT_TRUE(zip_archive_->fileExists(path));
    }
}

// 测试5: 大文件处理
TEST_F(ZipArchiveTest, LargeFileHandling) {
    LOG_INFO("Testing large file handling");

    // 创建 ZIP 文件
    zip_archive_ = std::make_unique<ZipArchive>(test_zip_path_);
    ASSERT_TRUE(zip_archive_->open(true));
    
    // 创建大文件数据 (1MB)
    const size_t large_size = 1024 * 1024;
    std::string large_content = createTestString(large_size);
    std::string internal_path = "large_file.txt";
    
    // 添加大文件到 ZIP
    EXPECT_TRUE(zip_archive_->addFile(internal_path, large_content));
    
    // 关闭 ZIP 文件
    zip_archive_->close();
    
    // 确保文件已经完全写入磁盘
    std::this_thread::sleep_for(std::chrono::milliseconds(10));
    
    // 重新打开 ZIP 文件进行读取
    ASSERT_TRUE(zip_archive_->open(false));
    
    // 检查文件是否存在
    EXPECT_TRUE(zip_archive_->fileExists(internal_path));
    
    // 提取文件内容
    std::string extracted_content;
    EXPECT_TRUE(zip_archive_->extractFile(internal_path, extracted_content));
    
    // 验证内容是否一致
    EXPECT_EQ(extracted_content.size(), large_content.size());
    EXPECT_EQ(extracted_content, large_content);
}

// 测试6: 特殊字符文件名
TEST_F(ZipArchiveTest, SpecialCharacterFilename) {
    LOG_INFO("Testing special characters in filenames");

    // 创建 ZIP 文件
    zip_archive_ = std::make_unique<ZipArchive>(test_zip_path_);
    ASSERT_TRUE(zip_archive_->open(true));
    
    // 创建包含特殊字符的文件名
    std::vector<std::string> special_filenames = {
        "file with spaces.txt",
        "file-with-dashes.txt",
        "file_with_underscores.txt",
        "file.with.dots.txt",
        "file@with#special$chars.txt",
        "中文文件名.txt",
        "file with ümläuts.txt"
    };
    
    // 添加文件
    for (const auto& filename : special_filenames) {
        std::string content = "Content for " + filename;
        EXPECT_TRUE(zip_archive_->addFile(filename, content));
    }
    
    // 关闭 ZIP 文件
    zip_archive_->close();
    
    // 确保文件已经完全写入磁盘
    std::this_thread::sleep_for(std::chrono::milliseconds(10));
    
    // 重新打开 ZIP 文件进行读取
    ASSERT_TRUE(zip_archive_->open(false));
    
    // 验证所有文件都存在
    for (const auto& filename : special_filenames) {
        EXPECT_TRUE(zip_archive_->fileExists(filename));
        
        // 提取并验证内容
        std::string extracted_content;
        EXPECT_TRUE(zip_archive_->extractFile(filename, extracted_content));
        EXPECT_EQ(extracted_content, "Content for " + filename);
    }
}

// 测试7: 空文件处理
TEST_F(ZipArchiveTest, EmptyFileHandling) {
    LOG_INFO("Testing empty file handling");

    // 创建 ZIP 文件
    zip_archive_ = std::make_unique<ZipArchive>(test_zip_path_);
    ASSERT_TRUE(zip_archive_->open(true));
    
    // 添加空文件
    std::string empty_content = "";
    std::string internal_path = "empty_file.txt";
    
    EXPECT_TRUE(zip_archive_->addFile(internal_path, empty_content));
    
    // 关闭 ZIP 文件
    zip_archive_->close();
    
    // 确保文件已经完全写入磁盘
    std::this_thread::sleep_for(std::chrono::milliseconds(10));
    
    // 重新打开 ZIP 文件进行读取
    ASSERT_TRUE(zip_archive_->open(false));
    
    // 检查文件是否存在
    EXPECT_TRUE(zip_archive_->fileExists(internal_path));
    
    // 提取文件内容
    std::string extracted_content;
    EXPECT_TRUE(zip_archive_->extractFile(internal_path, extracted_content));
    
    // 验证内容是否为空
    EXPECT_EQ(extracted_content, empty_content);
}

// 测试8: 文件不存在的情况
TEST_F(ZipArchiveTest, NonExistentFile) {
    LOG_INFO("Testing non-existent file handling");

    // 创建 ZIP 文件
    zip_archive_ = std::make_unique<ZipArchive>(test_zip_path_);
    ASSERT_TRUE(zip_archive_->open(true));
    
    // 添加一个文件
    std::string internal_path = "existing_file.txt";
    std::string content = "This file exists";
    EXPECT_TRUE(zip_archive_->addFile(internal_path, content));
    
    // 关闭 ZIP 文件
    zip_archive_->close();
    
    // 确保文件已经完全写入磁盘
    std::this_thread::sleep_for(std::chrono::milliseconds(10));
    
    // 重新打开 ZIP 文件进行读取
    ASSERT_TRUE(zip_archive_->open(false));
    
    // 检查存在的文件
    EXPECT_TRUE(zip_archive_->fileExists(internal_path));
    
    // 检查不存在的文件
    EXPECT_FALSE(zip_archive_->fileExists("non_existent_file.txt"));
    
    // 尝试提取不存在的文件
    std::string extracted_content;
    EXPECT_FALSE(zip_archive_->extractFile("non_existent_file.txt", extracted_content));
}

// 测试9: 重复添加同名文件
TEST_F(ZipArchiveTest, DuplicateFilename) {
    LOG_INFO("Testing duplicate filename handling");

    // 创建 ZIP 文件
    zip_archive_ = std::make_unique<ZipArchive>(test_zip_path_);
    ASSERT_TRUE(zip_archive_->open(true));
    
    // 添加文件
    std::string internal_path = "duplicate.txt";
    std::string first_content = "First content";
    std::string second_content = "Second content";
    
    EXPECT_TRUE(zip_archive_->addFile(internal_path, first_content));
    
    // 再次添加同名文件（应该覆盖）
    EXPECT_TRUE(zip_archive_->addFile(internal_path, second_content));
    
    // 关闭 ZIP 文件
    zip_archive_->close();
    
    // 确保文件已经完全写入磁盘
    std::this_thread::sleep_for(std::chrono::milliseconds(10));
    
    // 重新打开 ZIP 文件进行读取
    ASSERT_TRUE(zip_archive_->open(false));
    
    // 检查文件是否存在
    EXPECT_TRUE(zip_archive_->fileExists(internal_path));
    
    // 提取文件内容
    std::string extracted_content;
    EXPECT_TRUE(zip_archive_->extractFile(internal_path, extracted_content));
    
    // 验证内容是第二次添加的内容
    EXPECT_EQ(extracted_content, second_content);
}

// 测试10: 目录结构处理
TEST_F(ZipArchiveTest, DirectoryStructure) {
    LOG_INFO("Testing directory structure handling");

    // 创建 ZIP 文件
    zip_archive_ = std::make_unique<ZipArchive>(test_zip_path_);
    ASSERT_TRUE(zip_archive_->open(true));
    
    // 添加不同目录的文件
    std::vector<std::pair<std::string, std::string>> files = {
        {"root_file.txt", "Root file content"},
        {"dir1/subdir1/file1.txt", "File in subdir1"},
        {"dir1/subdir2/file2.txt", "File in subdir2"},
        {"dir2/file3.txt", "File in dir2"},
        {"dir1/subdir1/subsubdir/file4.txt", "File in subsubdir"}
    };
    
    for (const auto& [path, content] : files) {
        EXPECT_TRUE(zip_archive_->addFile(path, content));
    }
    
    // 关闭 ZIP 文件
    zip_archive_->close();
    
    // 确保文件已经完全写入磁盘
    std::this_thread::sleep_for(std::chrono::milliseconds(10));
    
    // 重新打开 ZIP 文件进行读取
    ASSERT_TRUE(zip_archive_->open(false));
    
    // 验证所有文件都存在
    for (const auto& [path, content] : files) {
        EXPECT_TRUE(zip_archive_->fileExists(path));
        
        // 提取并验证内容
        std::string extracted_content;
        EXPECT_TRUE(zip_archive_->extractFile(path, extracted_content));
        EXPECT_EQ(extracted_content, content);
    }
    
    // 列出所有文件
    std::vector<std::string> file_list = zip_archive_->listFiles();
    EXPECT_EQ(file_list.size(), files.size());
}

}} // namespace fastexcel::archive