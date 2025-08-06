#include <gtest/gtest.h>
#include "fastexcel/FastExcel.hpp"
#include <filesystem>
#include <fstream>

class ZipCompatibilityTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 初始化FastExcel库
        ASSERT_TRUE(fastexcel::initialize());
    }
    
    void TearDown() override {
        // 清理FastExcel库
        fastexcel::cleanup();
        
        // 清理测试文件
        cleanupTestFiles();
    }
    
    void cleanupTestFiles() {
        std::vector<std::string> test_files = {
            "test_auto_compatibility.xlsx",
            "test_batch_compatibility.xlsx", 
            "test_streaming_compatibility.xlsx"
        };
        
        for (const auto& file : test_files) {
            if (std::filesystem::exists(file)) {
                std::filesystem::remove(file);
            }
        }
    }
    
    bool validateZipFile(const std::string& filename) {
        std::ifstream file(filename, std::ios::binary);
        if (!file.is_open()) {
            return false;
        }
        
        // 检查ZIP文件头 (PK\x03\x04)
        char header[4];
        file.read(header, 4);
        return (header[0] == 'P' && header[1] == 'K' && 
                header[2] == 0x03 && header[3] == 0x04);
    }
    
    void testWorkbookMode(fastexcel::core::WorkbookMode mode, 
                         const std::string& filename, 
                         const std::string& modeName) {
        // 创建工作簿并设置模式
        auto workbook = std::make_shared<fastexcel::core::Workbook>(filename);
        workbook->setMode(mode);
        
        ASSERT_TRUE(workbook->open()) << "Failed to open workbook in " << modeName << " mode";
        
        // 添加工作表
        auto worksheet = workbook->addWorksheet("TestSheet");
        ASSERT_NE(worksheet, nullptr) << "Failed to create worksheet in " << modeName << " mode";
        
        // 写入测试数据
        worksheet->writeString(0, 0, "Mode");
        worksheet->writeString(0, 1, modeName);
        worksheet->writeString(1, 0, "Test Data");
        worksheet->writeNumber(1, 1, 123.45);
        worksheet->writeString(2, 0, "Excel Compatibility");
        worksheet->writeString(2, 1, "PASSED");
        
        // 添加一些格式化数据
        for (int row = 4; row < 10; ++row) {
            worksheet->writeString(row, 0, "Row " + std::to_string(row + 1));
            worksheet->writeNumber(row, 1, row * 10.5);
            worksheet->writeString(row, 2, "Data " + std::to_string(row));
        }
        
        // 保存文件
        ASSERT_TRUE(workbook->save()) << "Failed to save workbook in " << modeName << " mode";
        workbook->close();
        
        // 验证文件存在
        ASSERT_TRUE(std::filesystem::exists(filename)) << "File not created in " << modeName << " mode";
        
        // 验证文件大小合理
        auto fileSize = std::filesystem::file_size(filename);
        EXPECT_GT(fileSize, 1000) << "File size too small in " << modeName << " mode";
        EXPECT_LT(fileSize, 1024 * 1024) << "File size too large in " << modeName << " mode";
        
        // 验证ZIP文件结构
        EXPECT_TRUE(validateZipFile(filename)) << "Invalid ZIP structure in " << modeName << " mode";
    }
};

TEST_F(ZipCompatibilityTest, AutoModeCompatibility) {
    testWorkbookMode(fastexcel::core::WorkbookMode::AUTO, 
                    "test_auto_compatibility.xlsx", "AUTO");
}

TEST_F(ZipCompatibilityTest, BatchModeCompatibility) {
    testWorkbookMode(fastexcel::core::WorkbookMode::BATCH, 
                    "test_batch_compatibility.xlsx", "BATCH");
}

TEST_F(ZipCompatibilityTest, StreamingModeCompatibility) {
    testWorkbookMode(fastexcel::core::WorkbookMode::STREAMING, 
                    "test_streaming_compatibility.xlsx", "STREAMING");
}

TEST_F(ZipCompatibilityTest, ZipFileHeaderValidation) {
    // 创建一个简单的Excel文件
    auto workbook = std::make_shared<fastexcel::core::Workbook>("test_header.xlsx");
    ASSERT_TRUE(workbook->open());
    
    auto worksheet = workbook->addWorksheet("HeaderTest");
    worksheet->writeString(0, 0, "Header Test");
    worksheet->writeNumber(0, 1, 42);
    
    ASSERT_TRUE(workbook->save());
    workbook->close();
    
    // 验证ZIP文件头
    EXPECT_TRUE(validateZipFile("test_header.xlsx"));
    
    // 验证文件可以被重新读取
    std::ifstream file("test_header.xlsx", std::ios::binary);
    ASSERT_TRUE(file.is_open());
    
    // 读取更多字节以验证ZIP结构
    char buffer[30];
    file.read(buffer, 30);
    
    // 验证本地文件头签名
    EXPECT_EQ(buffer[0], 'P');
    EXPECT_EQ(buffer[1], 'K');
    EXPECT_EQ(buffer[2], 0x03);
    EXPECT_EQ(buffer[3], 0x04);
    
    file.close();
    
    // 清理测试文件
    std::filesystem::remove("test_header.xlsx");
}

TEST_F(ZipCompatibilityTest, MultipleWorksheetCompatibility) {
    auto workbook = std::make_shared<fastexcel::core::Workbook>("test_multiple_sheets.xlsx");
    ASSERT_TRUE(workbook->open());
    
    // 创建多个工作表
    auto sheet1 = workbook->addWorksheet("Sheet1");
    auto sheet2 = workbook->addWorksheet("Sheet2");
    auto sheet3 = workbook->addWorksheet("Sheet3");
    
    ASSERT_NE(sheet1, nullptr);
    ASSERT_NE(sheet2, nullptr);
    ASSERT_NE(sheet3, nullptr);
    
    // 在每个工作表中写入数据
    sheet1->writeString(0, 0, "This is Sheet 1");
    sheet1->writeNumber(1, 0, 100);
    
    sheet2->writeString(0, 0, "This is Sheet 2");
    sheet2->writeNumber(1, 0, 200);
    
    sheet3->writeString(0, 0, "This is Sheet 3");
    sheet3->writeNumber(1, 0, 300);
    
    ASSERT_TRUE(workbook->save());
    workbook->close();
    
    // 验证文件
    EXPECT_TRUE(std::filesystem::exists("test_multiple_sheets.xlsx"));
    EXPECT_TRUE(validateZipFile("test_multiple_sheets.xlsx"));
    
    // 清理测试文件
    std::filesystem::remove("test_multiple_sheets.xlsx");
}

TEST_F(ZipCompatibilityTest, LargeDataCompatibility) {
    auto workbook = std::make_shared<fastexcel::core::Workbook>("test_large_data.xlsx");
    ASSERT_TRUE(workbook->open());
    
    auto worksheet = workbook->addWorksheet("LargeData");
    ASSERT_NE(worksheet, nullptr);
    
    // 写入较大量的数据
    const int rows = 1000;
    const int cols = 10;
    
    for (int row = 0; row < rows; ++row) {
        for (int col = 0; col < cols; ++col) {
            if (col % 2 == 0) {
                worksheet->writeString(row, col, "Data_" + std::to_string(row) + "_" + std::to_string(col));
            } else {
                worksheet->writeNumber(row, col, row * col + 0.5);
            }
        }
    }
    
    ASSERT_TRUE(workbook->save());
    workbook->close();
    
    // 验证文件
    EXPECT_TRUE(std::filesystem::exists("test_large_data.xlsx"));
    EXPECT_TRUE(validateZipFile("test_large_data.xlsx"));
    
    // 验证文件大小合理
    auto fileSize = std::filesystem::file_size("test_large_data.xlsx");
    EXPECT_GT(fileSize, 10000) << "Large data file should be reasonably sized";
    
    // 清理测试文件
    std::filesystem::remove("test_large_data.xlsx");
}