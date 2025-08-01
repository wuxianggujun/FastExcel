#include <gtest/gtest.h>
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/Format.hpp"
#include <memory>
#include <filesystem>
#include <thread>
#include <atomic>

using namespace fastexcel::core;

class WorkbookTest : public ::testing::Test {
protected:
    void SetUp() override {
        test_filename = "test_workbook.xlsx";
        workbook = std::make_unique<Workbook>(test_filename);
    }
    
    void TearDown() override {
        workbook.reset();
        // 清理测试文件
        if (std::filesystem::exists(test_filename)) {
            std::filesystem::remove(test_filename);
        }
    }
    
    std::string test_filename;
    std::unique_ptr<Workbook> workbook;
};

// 测试工作簿创建
TEST_F(WorkbookTest, Creation) {
    EXPECT_NE(workbook, nullptr);
    EXPECT_EQ(workbook->getFilename(), test_filename);
    EXPECT_EQ(workbook->getWorksheetCount(), 0);
}

// 测试工作簿打开和关闭
TEST_F(WorkbookTest, OpenClose) {
    EXPECT_TRUE(workbook->open());
    EXPECT_TRUE(workbook->isOpen());
    
    EXPECT_TRUE(workbook->close());
    EXPECT_FALSE(workbook->isOpen());
}

// 测试添加工作表
TEST_F(WorkbookTest, AddWorksheet) {
    workbook->open();
    
    // 添加默认名称的工作表
    auto worksheet1 = workbook->addWorksheet();
    EXPECT_NE(worksheet1, nullptr);
    EXPECT_EQ(workbook->getWorksheetCount(), 1);
    EXPECT_EQ(worksheet1->getName(), "Sheet1");
    
    // 添加指定名称的工作表
    auto worksheet2 = workbook->addWorksheet("CustomSheet");
    EXPECT_NE(worksheet2, nullptr);
    EXPECT_EQ(workbook->getWorksheetCount(), 2);
    EXPECT_EQ(worksheet2->getName(), "CustomSheet");
    
    // 添加更多工作表
    auto worksheet3 = workbook->addWorksheet();
    EXPECT_EQ(worksheet3->getName(), "Sheet2");
    EXPECT_EQ(workbook->getWorksheetCount(), 3);
}

// 测试获取工作表
TEST_F(WorkbookTest, GetWorksheet) {
    workbook->open();
    
    auto worksheet1 = workbook->addWorksheet("First");
    auto worksheet2 = workbook->addWorksheet("Second");
    
    // 按名称获取
    auto retrieved1 = workbook->getWorksheet("First");
    EXPECT_EQ(retrieved1, worksheet1);
    
    auto retrieved2 = workbook->getWorksheet("Second");
    EXPECT_EQ(retrieved2, worksheet2);
    
    // 获取不存在的工作表
    auto nonexistent = workbook->getWorksheet("NonExistent");
    EXPECT_EQ(nonexistent, nullptr);
    
    // 按索引获取
    auto byIndex0 = workbook->getWorksheet(0);
    EXPECT_EQ(byIndex0, worksheet1);
    
    auto byIndex1 = workbook->getWorksheet(1);
    EXPECT_EQ(byIndex1, worksheet2);
    
    // 获取超出范围的索引
    auto outOfRange = workbook->getWorksheet(10);
    EXPECT_EQ(outOfRange, nullptr);
}

// 测试格式创建
TEST_F(WorkbookTest, CreateFormat) {
    workbook->open();
    
    auto format1 = workbook->createFormat();
    EXPECT_NE(format1, nullptr);
    EXPECT_EQ(format1->getXfIndex(), 0);
    
    auto format2 = workbook->createFormat();
    EXPECT_NE(format2, nullptr);
    EXPECT_EQ(format2->getXfIndex(), 1);
    
    // 验证格式是不同的对象
    EXPECT_NE(format1, format2);
}

// 测试获取格式
TEST_F(WorkbookTest, GetFormat) {
    workbook->open();
    
    auto format1 = workbook->createFormat();
    auto format2 = workbook->createFormat();
    
    // 按ID获取格式
    auto retrieved1 = workbook->getFormat(0);
    EXPECT_EQ(retrieved1, format1);
    
    auto retrieved2 = workbook->getFormat(1);
    EXPECT_EQ(retrieved2, format2);
    
    // 获取不存在的格式
    auto nonexistent = workbook->getFormat(100);
    EXPECT_EQ(nonexistent, nullptr);
}

// 测试文档属性
TEST_F(WorkbookTest, DocumentProperties) {
    workbook->open();
    
    // 设置基本属性
    workbook->setTitle("Test Title");
    EXPECT_EQ(workbook->getDocumentProperties().title, "Test Title");
    
    workbook->setAuthor("Test Author");
    EXPECT_EQ(workbook->getDocumentProperties().author, "Test Author");
    
    workbook->setSubject("Test Subject");
    EXPECT_EQ(workbook->getDocumentProperties().subject, "Test Subject");
    
    workbook->setKeywords("test, keywords");
    EXPECT_EQ(workbook->getDocumentProperties().keywords, "test, keywords");
    
    workbook->setComments("Test Comments");
    EXPECT_EQ(workbook->getDocumentProperties().comments, "Test Comments");
    
    workbook->setCompany("Test Company");
    EXPECT_EQ(workbook->getDocumentProperties().company, "Test Company");
    
    workbook->setManager("Test Manager");
    EXPECT_EQ(workbook->getDocumentProperties().manager, "Test Manager");
    
    workbook->setCategory("Test Category");
    EXPECT_EQ(workbook->getDocumentProperties().category, "Test Category");
}

// 测试自定义属性
TEST_F(WorkbookTest, CustomProperties) {
    workbook->open();
    
    // 分别测试每种类型，避免相互干扰
    
    // 测试字符串属性
    {
        std::string test_value = "Test Value";
        workbook->setCustomProperty("StringProp", test_value);
        std::string retrieved = workbook->getCustomProperty("StringProp");
        EXPECT_EQ(retrieved, "Test Value");
    }
    
    // 测试数字属性
    {
        workbook->setCustomProperty("NumberProp", 123.456);
        EXPECT_EQ(workbook->getCustomProperty("NumberProp"), "123.456000");
    }
    
    // 测试布尔属性
    {
        workbook->setCustomProperty("BoolProp", true);
        EXPECT_EQ(workbook->getCustomProperty("BoolProp"), "true");
        
        workbook->setCustomProperty("BoolProp2", false);
        EXPECT_EQ(workbook->getCustomProperty("BoolProp2"), "false");
    }
    
    // 最后再次验证字符串属性
    {
        std::string final_check = workbook->getCustomProperty("StringProp");
        EXPECT_EQ(final_check, "Test Value");
    }
}

// 测试定义名称
TEST_F(WorkbookTest, DefinedNames) {
    workbook->open();
    
    // 定义命名范围
    workbook->defineName("TestRange", "Sheet1!$A$1:$C$10");
    
    // 验证定义的名称
    EXPECT_EQ(workbook->getDefinedName("TestRange"), "Sheet1!$A$1:$C$10");
    
    // 定义更多名称
    workbook->defineName("AnotherRange", "Sheet2!$B$5:$D$15");
    EXPECT_EQ(workbook->getDefinedName("AnotherRange"), "Sheet2!$B$5:$D$15");
}

// 测试VBA项目
TEST_F(WorkbookTest, VBAProject) {
    workbook->open();
    
    std::string vba_path = "C:/Users/wuxianggujun/CodeSpace/CMakeProjects/FastExcel/test_vba.bin";
    workbook->addVbaProject(vba_path);
    
    EXPECT_TRUE(workbook->hasVbaProject());
}

// 测试常量内存模式
TEST_F(WorkbookTest, ConstantMemoryMode) {
    workbook->open();
    
    // 设置常量内存模式
    workbook->setConstantMemory(true);
    workbook->setConstantMemory(false);
    
    // 这个测试主要验证方法调用不会崩溃
    EXPECT_TRUE(true);
}

// 测试基本功能
TEST_F(WorkbookTest, BasicFunctionality) {
    workbook->open();
    
    // 添加一些内容
    auto worksheet = workbook->addWorksheet("TestSheet");
    auto format = workbook->createFormat();
    format->setBold(true);
    
    // 验证基本功能正常工作
    EXPECT_NE(worksheet, nullptr);
    EXPECT_NE(format, nullptr);
    EXPECT_EQ(worksheet->getName(), "TestSheet");
}

// 测试保存功能
TEST_F(WorkbookTest, Save) {
    workbook->open();
    
    // 添加一些内容
    auto worksheet = workbook->addWorksheet("TestSheet");
    worksheet->writeString(0, 0, "Hello");
    worksheet->writeNumber(0, 1, 123.45);
    
    // 保存文件
    EXPECT_TRUE(workbook->save());
    
    // 确保文件被完全关闭和刷新
    workbook->close();
    
    // 验证文件是否存在
    EXPECT_TRUE(std::filesystem::exists(test_filename));
    
    // 验证文件大小大于0
    auto file_size = std::filesystem::file_size(test_filename);
    EXPECT_GT(file_size, 0);
}

// 测试错误处理
TEST_F(WorkbookTest, ErrorHandling) {
    // 测试未打开工作簿时的操作
    EXPECT_THROW(workbook->addWorksheet(), std::runtime_error);
    EXPECT_THROW(workbook->createFormat(), std::runtime_error);
    EXPECT_THROW(workbook->save(), std::runtime_error);
    
    // 打开后应该正常
    workbook->open();
    EXPECT_NO_THROW(workbook->addWorksheet());
    EXPECT_NO_THROW(workbook->createFormat());
}

// 测试重复名称的工作表
TEST_F(WorkbookTest, DuplicateWorksheetNames) {
    workbook->open();
    
    auto worksheet1 = workbook->addWorksheet("TestSheet");
    EXPECT_EQ(worksheet1->getName(), "TestSheet");
    
    // 添加重复名称的工作表，应该自动重命名
    auto worksheet2 = workbook->addWorksheet("TestSheet");
    EXPECT_NE(worksheet2->getName(), "TestSheet");
    EXPECT_EQ(worksheet2->getName(), "TestSheet1"); // 或类似的重命名策略
}

// 测试大量工作表
TEST_F(WorkbookTest, ManyWorksheets) {
    workbook->open();
    
    const int num_sheets = 10;
    std::vector<std::shared_ptr<Worksheet>> sheets;
    
    for (int i = 0; i < num_sheets; ++i) {
        auto sheet = workbook->addWorksheet("Sheet" + std::to_string(i));
        sheets.push_back(sheet);
    }
    
    EXPECT_EQ(workbook->getWorksheetCount(), num_sheets);
    
    // 验证所有工作表都能正确获取
    for (int i = 0; i < num_sheets; ++i) {
        auto retrieved = workbook->getWorksheet(i);
        EXPECT_EQ(retrieved, sheets[i]);
    }
}

// 测试大量格式
TEST_F(WorkbookTest, ManyFormats) {
    workbook->open();
    
    const int num_formats = 100;
    std::vector<std::shared_ptr<Format>> formats;
    
    for (int i = 0; i < num_formats; ++i) {
        auto format = workbook->createFormat();
        format->setBold(i % 2 == 0);
        format->setItalic(i % 3 == 0);
        formats.push_back(format);
    }
    
    // 验证所有格式都能正确获取
    for (int i = 0; i < num_formats; ++i) {
        auto retrieved = workbook->getFormat(i);
        EXPECT_EQ(retrieved, formats[i]);
        EXPECT_EQ(retrieved->getXfIndex(), i);
    }
}

// 测试内存管理
TEST_F(WorkbookTest, MemoryManagement) {
    workbook->open();
    
    // 创建大量对象
    std::vector<std::weak_ptr<Worksheet>> weak_sheets;
    std::vector<std::weak_ptr<Format>> weak_formats;
    
    {
        // 在作用域内创建对象
        for (int i = 0; i < 10; ++i) {
            auto sheet = workbook->addWorksheet("TempSheet" + std::to_string(i));
            auto format = workbook->createFormat();
            
            weak_sheets.push_back(sheet);
            weak_formats.push_back(format);
        }
    }
    
    // 对象应该仍然存在（被workbook持有）
    for (const auto& weak_sheet : weak_sheets) {
        EXPECT_FALSE(weak_sheet.expired());
    }
    
    for (const auto& weak_format : weak_formats) {
        EXPECT_FALSE(weak_format.expired());
    }
}

// 测试线程安全（基本测试）
TEST_F(WorkbookTest, ThreadSafety) {
    workbook->open();
    
    // 这是一个基本的线程安全测试
    // 在实际实现中，可能需要更复杂的并发测试
    std::vector<std::thread> threads;
    std::atomic<int> success_count(0);
    
    for (int i = 0; i < 5; ++i) {
        threads.emplace_back([this, i, &success_count]() {
            try {
                auto format = workbook->createFormat();
                format->setBold(true);
                success_count++;
            } catch (...) {
                // 线程安全问题可能导致异常
            }
        });
    }
    
    for (auto& thread : threads) {
        thread.join();
    }
    
    // 在线程安全的实现中，所有操作都应该成功
    EXPECT_EQ(success_count.load(), 5);
}