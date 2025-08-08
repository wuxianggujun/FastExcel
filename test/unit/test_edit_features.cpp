/**
 * @file test_edit_features.cpp
 * @brief FastExcel编辑功能单元测试
 */

#include <gtest/gtest.h>
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/reader/XLSXReader.hpp"
#include <fstream>
#include <memory>

using namespace fastexcel;

class EditFeaturesTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 创建测试工作簿
        test_workbook = core::Workbook::create("test_edit.xlsx");
        ASSERT_TRUE(test_workbook->open());
        
        test_worksheet = test_workbook->addWorksheet("TestSheet");
        ASSERT_NE(test_worksheet, nullptr);
        
        // 添加测试数据
        test_worksheet->writeString(0, 0, "Name");
        test_worksheet->writeString(0, 1, "Age");
        test_worksheet->writeString(0, 2, "Department");
        
        test_worksheet->writeString(1, 0, "Alice");
        test_worksheet->writeNumber(1, 1, 25);
        test_worksheet->writeString(1, 2, "Engineering");
        
        test_worksheet->writeString(2, 0, "Bob");
        test_worksheet->writeNumber(2, 1, 30);
        test_worksheet->writeString(2, 2, "Sales");
        
        test_worksheet->writeString(3, 0, "Charlie");
        test_worksheet->writeNumber(3, 1, 28);
        test_worksheet->writeString(3, 2, "Engineering");
    }
    
    void TearDown() override {
        if (test_workbook) {
            test_workbook->close();
        }
        
        // 清理测试文件
        std::remove("test_edit.xlsx");
        std::remove("test_edit_copy.xlsx");
    }
    
    std::unique_ptr<core::Workbook> test_workbook;
    std::shared_ptr<core::Worksheet> test_worksheet;
};

// 测试单元格编辑功能
TEST_F(EditFeaturesTest, CellEditValue) {
    // 测试编辑字符串值
    test_worksheet->editCellValue(1, 0, "Alice Smith", true);
    const auto& cell = test_worksheet->getCell(1, 0);
    EXPECT_EQ(cell.getStringValue(), "Alice Smith");
    
    // 测试编辑数字值
    test_worksheet->editCellValue(1, 1, 26.0, true);
    const auto& age_cell = test_worksheet->getCell(1, 1);
    EXPECT_DOUBLE_EQ(age_cell.getNumberValue(), 26.0);
    
    // 测试编辑布尔值
    test_worksheet->writeBoolean(1, 3, true);
    test_worksheet->editCellValue(1, 3, false, true);
    const auto& bool_cell = test_worksheet->getCell(1, 3);
    EXPECT_FALSE(bool_cell.getBooleanValue());
}

// 测试单元格复制功能
TEST_F(EditFeaturesTest, CellCopy) {
    // 复制单个单元格
    test_worksheet->copyCell(1, 0, 4, 0, true);
    const auto& copied_cell = test_worksheet->getCell(4, 0);
    EXPECT_EQ(copied_cell.getStringValue(), "Alice");
    
    // 复制范围
    test_worksheet->copyRange(1, 0, 1, 2, 5, 0, true);
    EXPECT_EQ(test_worksheet->getCell(5, 0).getStringValue(), "Alice");
    EXPECT_DOUBLE_EQ(test_worksheet->getCell(5, 1).getNumberValue(), 25.0);
    EXPECT_EQ(test_worksheet->getCell(5, 2).getStringValue(), "Engineering");
}

// 测试单元格移动功能
TEST_F(EditFeaturesTest, CellMove) {
    // 移动单个单元格
    test_worksheet->moveCell(1, 0, 4, 0);
    
    // 检查源位置是否为空
    EXPECT_TRUE(test_worksheet->getCell(1, 0).isEmpty());
    
    // 检查目标位置是否有正确的值
    EXPECT_EQ(test_worksheet->getCell(4, 0).getStringValue(), "Alice");
}

// 测试查找和替换功能
TEST_F(EditFeaturesTest, FindAndReplace) {
    // 查找功能
    auto results = test_worksheet->findCells("Engineering", false, false);
    EXPECT_EQ(results.size(), 2); // Alice和Charlie都在Engineering部门
    
    // 替换功能
    int replacements = test_worksheet->findAndReplace("Engineering", "Development", false, false);
    EXPECT_EQ(replacements, 2);
    
    // 验证替换结果
    EXPECT_EQ(test_worksheet->getCell(1, 2).getStringValue(), "Development");
    EXPECT_EQ(test_worksheet->getCell(3, 2).getStringValue(), "Development");
}

// 测试排序功能
TEST_F(EditFeaturesTest, SortRange) {
    // 按年龄升序排序
    test_worksheet->sortRange(1, 0, 3, 2, 1, true, false);
    
    // 验证排序结果（按年龄：Alice 25, Charlie 28, Bob 30）
    EXPECT_EQ(test_worksheet->getCell(1, 0).getStringValue(), "Alice");
    EXPECT_EQ(test_worksheet->getCell(2, 0).getStringValue(), "Charlie");
    EXPECT_EQ(test_worksheet->getCell(3, 0).getStringValue(), "Bob");
}

// 测试工作簿级别的编辑功能
class WorkbookEditTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 创建测试工作簿
        workbook = core::Workbook::create("test_workbook_edit.xlsx");
        ASSERT_TRUE(workbook->open());
        
        // 添加多个工作表
        sheet1 = workbook->addWorksheet("Sheet1");
        sheet2 = workbook->addWorksheet("Sheet2");
        sheet3 = workbook->addWorksheet("Sheet3");
        
        // 在工作表中添加一些数据
        sheet1->writeString(0, 0, "Data in Sheet1");
        sheet2->writeString(0, 0, "Data in Sheet2");
        sheet3->writeString(0, 0, "Data in Sheet3");
        
        // 保存文件
        ASSERT_TRUE(workbook->save());
    }
    
    void TearDown() override {
        if (workbook) {
            workbook->close();
        }
        
        // 清理测试文件
        std::remove("test_workbook_edit.xlsx");
        std::remove("test_workbook_merged.xlsx");
        std::remove("test_workbook_export.xlsx");
    }
    
    std::unique_ptr<core::Workbook> workbook;
    std::shared_ptr<core::Worksheet> sheet1, sheet2, sheet3;
};

// 测试批量重命名工作表
TEST_F(WorkbookEditTest, BatchRenameWorksheets) {
    std::unordered_map<std::string, std::string> rename_map = {
        {"Sheet1", "Data"},
        {"Sheet2", "Analysis"},
        {"Sheet3", "Summary"}
    };
    
    int renamed = workbook->batchRenameWorksheets(rename_map);
    EXPECT_EQ(renamed, 3);
    
    // 验证重命名结果
    EXPECT_NE(workbook->getWorksheet("Data"), nullptr);
    EXPECT_NE(workbook->getWorksheet("Analysis"), nullptr);
    EXPECT_NE(workbook->getWorksheet("Summary"), nullptr);
    EXPECT_EQ(workbook->getWorksheet("Sheet1"), nullptr);
}

// 测试批量删除工作表
TEST_F(WorkbookEditTest, BatchRemoveWorksheets) {
    std::vector<std::string> to_remove = {"Sheet2", "Sheet3"};
    
    int removed = workbook->batchRemoveWorksheets(to_remove);
    EXPECT_EQ(removed, 2);
    EXPECT_EQ(workbook->getWorksheetCount(), 1);
    EXPECT_NE(workbook->getWorksheet("Sheet1"), nullptr);
}

// 测试工作表重新排序
TEST_F(WorkbookEditTest, ReorderWorksheets) {
    std::vector<std::string> new_order = {"Sheet3", "Sheet1", "Sheet2"};
    
    bool success = workbook->reorderWorksheets(new_order);
    EXPECT_TRUE(success);
    
    // 验证新顺序
    auto names = workbook->getWorksheetNames();
    EXPECT_EQ(names[0], "Sheet3");
    EXPECT_EQ(names[1], "Sheet1");
    EXPECT_EQ(names[2], "Sheet2");
}

// 测试全局查找和替换
TEST_F(WorkbookEditTest, GlobalFindAndReplace) {
    // 在所有工作表中添加相同的文本
    sheet1->writeString(1, 0, "Test Data");
    sheet2->writeString(1, 0, "Test Data");
    sheet3->writeString(1, 0, "Test Data");
    
    // 全局查找
    core::Workbook::FindReplaceOptions options;
    auto results = workbook->findAll("Test Data", options);
    EXPECT_EQ(results.size(), 3);
    
    // 全局替换
    int replacements = workbook->findAndReplaceAll("Test Data", "Modified Data", options);
    EXPECT_EQ(replacements, 3);
    
    // 验证替换结果
    EXPECT_EQ(sheet1->getCell(1, 0).getStringValue(), "Modified Data");
    EXPECT_EQ(sheet2->getCell(1, 0).getStringValue(), "Modified Data");
    EXPECT_EQ(sheet3->getCell(1, 0).getStringValue(), "Modified Data");
}

// 测试工作簿统计信息
TEST_F(WorkbookEditTest, WorkbookStatistics) {
    auto stats = workbook->getStatistics();
    
    EXPECT_EQ(stats.total_worksheets, 3);
    EXPECT_GT(stats.total_cells, 0);
    EXPECT_GT(stats.memory_usage, 0);
    
    // 检查每个工作表的单元格数量
    EXPECT_GT(stats.worksheet_cell_counts["Sheet1"], 0);
    EXPECT_GT(stats.worksheet_cell_counts["Sheet2"], 0);
    EXPECT_GT(stats.worksheet_cell_counts["Sheet3"], 0);
}

// 测试文件加载和刷新功能
TEST_F(WorkbookEditTest, LoadForEditAndRefresh) {
    // 关闭当前工作簿
    workbook->close();
    
    // 使用open重新加载
    auto loaded_workbook = core::Workbook::open("test_workbook_edit.xlsx");
    ASSERT_NE(loaded_workbook, nullptr);
    
    // 验证加载的内容
    EXPECT_EQ(loaded_workbook->getWorksheetCount(), 3);
    auto loaded_sheet1 = loaded_workbook->getWorksheet("Sheet1");
    ASSERT_NE(loaded_sheet1, nullptr);
    EXPECT_EQ(loaded_sheet1->getCell(0, 0).getStringValue(), "Data in Sheet1");
    
    loaded_workbook->close();
}

// 性能测试
TEST(EditPerformanceTest, LargeDataEdit) {
    auto workbook = core::Workbook::create("performance_test.xlsx");
    ASSERT_TRUE(workbook->open());
    
    // 启用高性能模式
    workbook->setHighPerformanceMode(true);
    
    auto worksheet = workbook->addWorksheet("PerformanceTest");
    
    // 写入大量数据
    const int rows = 1000;
    const int cols = 10;
    
    auto start = std::chrono::high_resolution_clock::now();
    
    for (int row = 0; row < rows; ++row) {
        for (int col = 0; col < cols; ++col) {
            if (col % 2 == 0) {
                worksheet->writeString(row, col, "Text" + std::to_string(row * cols + col));
            } else {
                worksheet->writeNumber(row, col, static_cast<double>(row * cols + col));
            }
        }
    }
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);
    
    std::cout << "写入 " << rows * cols << " 个单元格耗时: " << duration.count() << "ms" << std::endl;
    
    // 测试批量编辑性能
    start = std::chrono::high_resolution_clock::now();
    
    int replacements = worksheet->findAndReplace("Text", "Modified", false, false);
    
    end = std::chrono::high_resolution_clock::now();
    duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);
    
    std::cout << "查找替换 " << replacements << " 个单元格耗时: " << duration.count() << "ms" << std::endl;
    
    EXPECT_GT(replacements, 0);
    
    workbook->close();
    std::remove("performance_test.xlsx");
}

int main(int argc, char** argv) {
    ::testing::InitGoogleTest(&argc, argv);
    return RUN_ALL_TESTS();
}