// FastExcel 库 - 高性能Excel文件处理库
// 组件：工作簿集成测试
//
// 本文件实现了工作簿相关类的集成测试，使用 Gtest 框架测试多个模块的交互。

#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/Cell.hpp"
#include <gtest/gtest.h>
#include <string>
#include <memory>

namespace fastexcel {
namespace core {

// 工作簿集成测试套件
class WorkbookIntegrationTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 测试前的设置
        workbook = Workbook::create("test_workbook.xlsx");
    }

    void TearDown() override {
        // 测试后的清理
        workbook.reset();
    }

    std::unique_ptr<Workbook> workbook;
};

// 测试工作簿创建和工作表添加
TEST_F(WorkbookIntegrationTest, CreateWorkbookAndAddWorksheet) {
    // 添加工作表
    auto worksheet = workbook->addSheet("Sheet1");
    
    EXPECT_NE(worksheet, nullptr);
    EXPECT_EQ(worksheet->getName(), "Sheet1");
    
    // 验证工作簿中的工作表数量
    EXPECT_GE(workbook->getSheetCount(), 1);
}

// 测试单元格操作
TEST_F(WorkbookIntegrationTest, CellOperations) {
    auto worksheet = workbook->addSheet("TestSheet");
    
    // 创建单元格并设置值
    worksheet->getCell(0, 0).setValue(std::string("Hello, World!")); // A1单元格
    
    EXPECT_EQ(worksheet->getCell(0, 0).getType(), CellType::String);
    EXPECT_EQ(worksheet->getCell(0, 0).getValue<std::string>(), "Hello, World!");
    
    // 测试数值单元格
    worksheet->getCell(0, 1).setValue(42.5); // B1单元格
    
    EXPECT_EQ(worksheet->getCell(0, 1).getType(), CellType::Number);
    EXPECT_DOUBLE_EQ(worksheet->getCell(0, 1).getValue<double>(), 42.5);
    
    // 测试布尔单元格
    worksheet->getCell(0, 2).setValue(true); // C1单元格
    
    EXPECT_EQ(worksheet->getCell(0, 2).getType(), CellType::Boolean);
    EXPECT_TRUE(worksheet->getCell(0, 2).getValue<bool>());
}

// 测试工作表间的数据复制
TEST_F(WorkbookIntegrationTest, WorksheetDataCopy) {
    auto sourceSheet = workbook->addSheet("Source");
    auto targetSheet = workbook->addSheet("Target");
    
    // 在源工作表中设置数据
    sourceSheet->getCell(0, 0).setValue(std::string("Copied Data"));
    
    // 复制数据到目标工作表
    targetSheet->getCell(0, 0).setValue(sourceSheet->getCell(0, 0).getValue<std::string>());
    
    EXPECT_EQ(targetSheet->getCell(0, 0).getValue<std::string>(), "Copied Data");
    EXPECT_EQ(targetSheet->getCell(0, 0).getType(), CellType::String);
}

// 测试公式单元格
TEST_F(WorkbookIntegrationTest, FormulaCell) {
    auto worksheet = workbook->addSheet("FormulaSheet");
    
    // 设置一些基础数据
    worksheet->getCell(0, 0).setValue(10); // A1
    worksheet->getCell(0, 1).setValue(20); // B1
    
    // 创建公式单元格
    worksheet->getCell(0, 2).setFormula("=A1+B1"); // C1
    
    EXPECT_EQ(worksheet->getCell(0, 2).getType(), CellType::Formula);
    EXPECT_EQ(worksheet->getCell(0, 2).getFormula(), "=A1+B1");
}

// 测试工作簿保存和加载
TEST_F(WorkbookIntegrationTest, SaveAndLoadWorkbook) {
    auto worksheet = workbook->addSheet("TestData");
    
    // 添加一些测试数据
    worksheet->getCell(0, 0).setValue(std::string("Test String"));
    worksheet->getCell(0, 1).setValue(123.45);
    worksheet->getCell(0, 2).setValue(true);
    
    // 保存工作簿
    EXPECT_TRUE(workbook->save());
    
    // 创建新工作簿并加载保存的文件
    auto loadedWorkbook = Workbook::openReadOnly("test_workbook.xlsx");
    EXPECT_NE(loadedWorkbook, nullptr);
    
    // 验证加载的数据
    auto loadedWorksheet = loadedWorkbook->getSheet(0);
    EXPECT_NE(loadedWorksheet, nullptr);
    
    EXPECT_EQ(loadedWorksheet->getCell(0, 0).getType(), CellType::String);
    EXPECT_EQ(loadedWorksheet->getCell(0, 0).getValue<std::string>(), "Test String");
}

// 测试多工作表操作
TEST_F(WorkbookIntegrationTest, MultipleWorksheets) {
    // 添加多个工作表
    auto sheet1 = workbook->addSheet("Sheet1");
    auto sheet2 = workbook->addSheet("Sheet2");
    auto sheet3 = workbook->addSheet("Sheet3");
    
    // 在不同工作表中设置数据
    sheet1->getCell(0, 0).setValue(std::string("Data from Sheet1"));
    sheet2->getCell(0, 0).setValue(std::string("Data from Sheet2"));
    sheet3->getCell(0, 0).setValue(std::string("Data from Sheet3"));
    
    // 验证数据
    EXPECT_EQ(sheet1->getCell(0, 0).getValue<std::string>(), "Data from Sheet1");
    EXPECT_EQ(sheet2->getCell(0, 0).getValue<std::string>(), "Data from Sheet2");
    EXPECT_EQ(sheet3->getCell(0, 0).getValue<std::string>(), "Data from Sheet3");
    
    // 验证工作表数量
    EXPECT_EQ(workbook->getSheetCount(), 3);
}

// 测试工作表重命名
TEST_F(WorkbookIntegrationTest, RenameWorksheet) {
    auto worksheet = workbook->addSheet("OriginalName");
    
    // 重命名工作表
    worksheet->setName("NewName");
    EXPECT_EQ(worksheet->getName(), "NewName");
}

// 测试单元格格式应用
TEST_F(WorkbookIntegrationTest, CellFormatting) {
    auto worksheet = workbook->addSheet("FormatTest");
    
    // 创建格式
    auto format = std::make_shared<FormatDescriptor>(StyleBuilder().bold().build());
    
    // 应用格式到单元格
    worksheet->getCell(0, 0).setFormat(format);
    EXPECT_EQ(worksheet->getCell(0, 0).getFormatDescriptor(), format);
}

}} // namespace fastexcel::core