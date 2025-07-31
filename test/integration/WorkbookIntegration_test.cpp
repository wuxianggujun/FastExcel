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
        workbook = std::make_unique<Workbook>("test_workbook.xlsx");
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
    auto worksheet = workbook->addWorksheet("Sheet1");
    
    EXPECT_NE(worksheet, nullptr);
    EXPECT_EQ(worksheet->getName(), "Sheet1");
    
    // 验证工作簿中的工作表数量
    EXPECT_GE(workbook->getWorksheetCount(), 1);
}

// 测试单元格操作
TEST_F(WorkbookIntegrationTest, CellOperations) {
    auto worksheet = workbook->addWorksheet("TestSheet");
    
    // 创建单元格并设置值
    auto cell = worksheet->getCell(0, 0); // A1单元格
    cell->setValue("Hello, World!");
    
    EXPECT_EQ(cell->getType(), CellType::String);
    EXPECT_EQ(cell->getStringValue(), "Hello, World!");
    
    // 测试数值单元格
    auto numberCell = worksheet->getCell(0, 1); // B1单元格
    numberCell->setValue(42.5);
    
    EXPECT_EQ(numberCell->getType(), CellType::Number);
    EXPECT_DOUBLE_EQ(numberCell->getNumberValue(), 42.5);
    
    // 测试布尔单元格
    auto boolCell = worksheet->getCell(0, 2); // C1单元格
    boolCell->setValue(true);
    
    EXPECT_EQ(boolCell->getType(), CellType::Boolean);
    EXPECT_TRUE(boolCell->getBooleanValue());
}

// 测试工作表间的数据复制
TEST_F(WorkbookIntegrationTest, WorksheetDataCopy) {
    auto sourceSheet = workbook->addWorksheet("Source");
    auto targetSheet = workbook->addWorksheet("Target");
    
    // 在源工作表中设置数据
    auto sourceCell = sourceSheet->getCell(0, 0);
    sourceCell->setValue("Copied Data");
    
    // 复制数据到目标工作表
    auto targetCell = targetSheet->getCell(0, 0);
    targetCell->setValue(sourceCell->getStringValue());
    
    EXPECT_EQ(targetCell->getStringValue(), "Copied Data");
    EXPECT_EQ(targetCell->getType(), CellType::String);
}

// 测试公式单元格
TEST_F(WorkbookIntegrationTest, FormulaCell) {
    auto worksheet = workbook->addWorksheet("FormulaSheet");
    
    // 设置一些基础数据
    worksheet->getCell(0, 0)->setValue(10); // A1
    worksheet->getCell(0, 1)->setValue(20); // B1
    
    // 创建公式单元格
    auto formulaCell = worksheet->getCell(0, 2); // C1
    formulaCell->setFormula("=A1+B1");
    
    EXPECT_EQ(formulaCell->getType(), CellType::Formula);
    EXPECT_EQ(formulaCell->getFormula(), "=A1+B1");
}

// 测试工作簿保存和加载
TEST_F(WorkbookIntegrationTest, SaveAndLoadWorkbook) {
    auto worksheet = workbook->addWorksheet("TestData");
    
    // 添加一些测试数据
    worksheet->getCell(0, 0)->setValue("Test String");
    worksheet->getCell(0, 1)->setValue(123.45);
    worksheet->getCell(0, 2)->setValue(true);
    
    // 保存工作簿
    EXPECT_TRUE(workbook->save());
    
    // 创建新工作簿并加载保存的文件
    auto loadedWorkbook = std::make_unique<Workbook>("test_workbook.xlsx");
    EXPECT_TRUE(loadedWorkbook->open());
    
    // 验证加载的数据
    auto loadedWorksheet = loadedWorkbook->getWorksheet(0);
    EXPECT_NE(loadedWorksheet, nullptr);
    
    auto loadedCell = loadedWorksheet->getCell(0, 0);
    EXPECT_EQ(loadedCell->getType(), CellType::String);
    EXPECT_EQ(loadedCell->getStringValue(), "Test String");
}

// 测试多工作表操作
TEST_F(WorkbookIntegrationTest, MultipleWorksheets) {
    // 添加多个工作表
    auto sheet1 = workbook->addWorksheet("Sheet1");
    auto sheet2 = workbook->addWorksheet("Sheet2");
    auto sheet3 = workbook->addWorksheet("Sheet3");
    
    // 在不同工作表中设置数据
    sheet1->getCell(0, 0)->setValue("Data from Sheet1");
    sheet2->getCell(0, 0)->setValue("Data from Sheet2");
    sheet3->getCell(0, 0)->setValue("Data from Sheet3");
    
    // 验证数据
    EXPECT_EQ(sheet1->getCell(0, 0)->getStringValue(), "Data from Sheet1");
    EXPECT_EQ(sheet2->getCell(0, 0)->getStringValue(), "Data from Sheet2");
    EXPECT_EQ(sheet3->getCell(0, 0)->getStringValue(), "Data from Sheet3");
    
    // 验证工作表数量
    EXPECT_EQ(workbook->getWorksheetCount(), 3);
}

// 测试工作表重命名
TEST_F(WorkbookIntegrationTest, RenameWorksheet) {
    auto worksheet = workbook->addWorksheet("OriginalName");
    
    // 重命名工作表
    EXPECT_TRUE(worksheet->rename("NewName"));
    EXPECT_EQ(worksheet->getName(), "NewName");
}

// 测试单元格格式应用
TEST_F(WorkbookIntegrationTest, CellFormatting) {
    auto worksheet = workbook->addWorksheet("FormatTest");
    auto cell = worksheet->getCell(0, 0);
    
    // 创建格式
    auto format = std::make_shared<Format>();
    // 这里可以设置具体的格式属性
    
    // 应用格式到单元格
    cell->setFormat(format);
    EXPECT_EQ(cell->getFormat(), format);
}

}} // namespace fastexcel::core