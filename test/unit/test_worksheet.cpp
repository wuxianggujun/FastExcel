#include <gtest/gtest.h>
#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/Format.hpp"
#include "fastexcel/core/Cell.hpp"
#include <memory>
#include <vector>
#include <chrono>

using namespace fastexcel::core;

class WorksheetTest : public ::testing::Test {
protected:
    void SetUp() override {
        workbook = Workbook::create("test.xlsx");
        workbook->open();
        worksheet = workbook->addWorksheet("TestSheet");
    }
    
    void TearDown() override {
        worksheet.reset();
        workbook.reset();
    }
    
    std::shared_ptr<Workbook> workbook;
    std::shared_ptr<Worksheet> worksheet;
};

// 测试工作表创建
TEST_F(WorksheetTest, Creation) {
    EXPECT_NE(worksheet, nullptr);
    EXPECT_EQ(worksheet->getName(), "TestSheet");
    EXPECT_EQ(worksheet->getUsedRange().first, -1);
    EXPECT_EQ(worksheet->getUsedRange().second, -1);
}

// 测试字符串写入
TEST_F(WorksheetTest, WriteString) {
    std::string test_value = "Hello, World!";
    worksheet->writeString(0, 0, test_value);
    
    const auto& cell = worksheet->getCell(0, 0);
    EXPECT_TRUE(cell.isString());
    EXPECT_EQ(cell.getStringValue(), test_value);
    
    auto [max_row, max_col] = worksheet->getUsedRange();
    EXPECT_EQ(max_row, 0);
    EXPECT_EQ(max_col, 0);
}

// 测试数字写入
TEST_F(WorksheetTest, WriteNumber) {
    double test_value = 123.456;
    worksheet->writeNumber(1, 1, test_value);
    
    const auto& cell = worksheet->getCell(1, 1);
    EXPECT_TRUE(cell.isNumber());
    EXPECT_DOUBLE_EQ(cell.getNumberValue(), test_value);
    
    auto [max_row, max_col] = worksheet->getUsedRange();
    EXPECT_EQ(max_row, 1);
    EXPECT_EQ(max_col, 1);
}

// 测试布尔值写入
TEST_F(WorksheetTest, WriteBoolean) {
    worksheet->writeBoolean(2, 2, true);
    worksheet->writeBoolean(2, 3, false);
    
    const auto& cell1 = worksheet->getCell(2, 2);
    EXPECT_TRUE(cell1.isBoolean());
    EXPECT_TRUE(cell1.getBooleanValue());
    
    const auto& cell2 = worksheet->getCell(2, 3);
    EXPECT_TRUE(cell2.isBoolean());
    EXPECT_FALSE(cell2.getBooleanValue());
}

// 测试公式写入
TEST_F(WorksheetTest, WriteFormula) {
    std::string formula = "SUM(A1:A10)";
    worksheet->writeFormula(3, 0, formula);
    
    const auto& cell = worksheet->getCell(3, 0);
    EXPECT_TRUE(cell.isFormula());
    EXPECT_EQ(cell.getFormula(), formula);
}

// 测试日期时间写入
TEST_F(WorksheetTest, WriteDateTime) {
    std::tm test_date = {};
    test_date.tm_year = 124; // 2024
    test_date.tm_mon = 0;    // January
    test_date.tm_mday = 1;   // 1st
    
    worksheet->writeDateTime(4, 0, test_date);
    
    const auto& cell = worksheet->getCell(4, 0);
    EXPECT_TRUE(cell.isNumber()); // 日期存储为数字
    EXPECT_GT(cell.getNumberValue(), 0);
}

// 测试超链接写入
TEST_F(WorksheetTest, WriteUrl) {
    std::string url = "https://www.example.com";
    std::string display_text = "Example";
    
    worksheet->writeUrl(5, 0, url, display_text);
    
    const auto& cell = worksheet->getCell(5, 0);
    EXPECT_TRUE(cell.isString());
    EXPECT_EQ(cell.getStringValue(), display_text);
    EXPECT_TRUE(cell.hasHyperlink());
    EXPECT_EQ(cell.getHyperlink(), url);
    
    // 测试没有显示文本的情况
    worksheet->writeUrl(5, 1, url);
    const auto& cell2 = worksheet->getCell(5, 1);
    EXPECT_EQ(cell2.getStringValue(), url);
}

// 测试格式应用
TEST_F(WorksheetTest, WriteWithFormat) {
    auto format = workbook->createFormat();
    format->setBold(true);
    format->setFontColor(COLOR_RED);
    
    worksheet->writeString(0, 0, "Formatted Text", format);
    
    const auto& cell = worksheet->getCell(0, 0);
    EXPECT_EQ(cell.getFormat(), format);
    EXPECT_TRUE(cell.getFormat()->isBold());
    EXPECT_EQ(cell.getFormat()->getFontColor(), COLOR_RED);
}

// 测试批量数据写入 - 字符串
TEST_F(WorksheetTest, WriteStringRange) {
    std::vector<std::vector<std::string>> data = {
        {"A1", "B1", "C1"},
        {"A2", "B2", "C2"},
        {"A3", "B3", "C3"}
    };
    
    worksheet->writeRange(0, 0, data);
    
    // 验证数据
    for (size_t row = 0; row < data.size(); ++row) {
        for (size_t col = 0; col < data[row].size(); ++col) {
            const auto& cell = worksheet->getCell(static_cast<int>(row), static_cast<int>(col));
            EXPECT_EQ(cell.getStringValue(), data[row][col]);
        }
    }
    
    auto [max_row, max_col] = worksheet->getUsedRange();
    EXPECT_EQ(max_row, 2);
    EXPECT_EQ(max_col, 2);
}

// 测试批量数据写入 - 数字
TEST_F(WorksheetTest, WriteNumberRange) {
    std::vector<std::vector<double>> data = {
        {1.1, 2.2, 3.3},
        {4.4, 5.5, 6.6}
    };
    
    worksheet->writeRange(0, 0, data);
    
    // 验证数据
    for (size_t row = 0; row < data.size(); ++row) {
        for (size_t col = 0; col < data[row].size(); ++col) {
            const auto& cell = worksheet->getCell(static_cast<int>(row), static_cast<int>(col));
            EXPECT_DOUBLE_EQ(cell.getNumberValue(), data[row][col]);
        }
    }
}

// 测试列宽设置
TEST_F(WorksheetTest, ColumnWidth) {
    // 设置单列宽度
    worksheet->setColumnWidth(0, 15.0);
    EXPECT_DOUBLE_EQ(worksheet->getColumnWidth(0), 15.0);
    
    // 设置多列宽度
    worksheet->setColumnWidth(1, 3, 20.0);
    EXPECT_DOUBLE_EQ(worksheet->getColumnWidth(1), 20.0);
    EXPECT_DOUBLE_EQ(worksheet->getColumnWidth(2), 20.0);
    EXPECT_DOUBLE_EQ(worksheet->getColumnWidth(3), 20.0);
}

// 测试行高设置
TEST_F(WorksheetTest, RowHeight) {
    worksheet->setRowHeight(0, 25.0);
    EXPECT_DOUBLE_EQ(worksheet->getRowHeight(0), 25.0);
}

// 测试列格式设置
TEST_F(WorksheetTest, ColumnFormat) {
    auto format = workbook->createFormat();
    format->setBold(true);
    
    worksheet->setColumnFormat(0, format);
    EXPECT_EQ(worksheet->getColumnFormat(0), format);
    
    // 设置多列格式
    worksheet->setColumnFormat(1, 3, format);
    EXPECT_EQ(worksheet->getColumnFormat(1), format);
    EXPECT_EQ(worksheet->getColumnFormat(2), format);
    EXPECT_EQ(worksheet->getColumnFormat(3), format);
}

// 测试行格式设置
TEST_F(WorksheetTest, RowFormat) {
    auto format = workbook->createFormat();
    format->setItalic(true);
    
    worksheet->setRowFormat(0, format);
    EXPECT_EQ(worksheet->getRowFormat(0), format);
}

// 测试隐藏列
TEST_F(WorksheetTest, HideColumn) {
    worksheet->hideColumn(0);
    EXPECT_TRUE(worksheet->isColumnHidden(0));
    
    // 隐藏多列
    worksheet->hideColumn(1, 3);
    EXPECT_TRUE(worksheet->isColumnHidden(1));
    EXPECT_TRUE(worksheet->isColumnHidden(2));
    EXPECT_TRUE(worksheet->isColumnHidden(3));
}

// 测试隐藏行
TEST_F(WorksheetTest, HideRow) {
    worksheet->hideRow(0);
    EXPECT_TRUE(worksheet->isRowHidden(0));
    
    // 隐藏多行
    worksheet->hideRow(1, 3);
    EXPECT_TRUE(worksheet->isRowHidden(1));
    EXPECT_TRUE(worksheet->isRowHidden(2));
    EXPECT_TRUE(worksheet->isRowHidden(3));
}

// 测试合并单元格
TEST_F(WorksheetTest, MergeCells) {
    worksheet->mergeCells(0, 0, 2, 2);
    
    auto merge_ranges = worksheet->getMergeRanges();
    EXPECT_EQ(merge_ranges.size(), 1);
    EXPECT_EQ(merge_ranges[0].first_row, 0);
    EXPECT_EQ(merge_ranges[0].first_col, 0);
    EXPECT_EQ(merge_ranges[0].last_row, 2);
    EXPECT_EQ(merge_ranges[0].last_col, 2);
}

// 测试带值的合并单元格
TEST_F(WorksheetTest, MergeRange) {
    auto format = workbook->createFormat();
    format->setHorizontalAlign(HorizontalAlign::Center);
    
    worksheet->mergeRange(0, 0, 0, 3, "Merged Title", format);
    
    // 验证合并范围
    auto merge_ranges = worksheet->getMergeRanges();
    EXPECT_EQ(merge_ranges.size(), 1);
    
    // 验证单元格内容
    const auto& cell = worksheet->getCell(0, 0);
    EXPECT_EQ(cell.getStringValue(), "Merged Title");
    EXPECT_EQ(cell.getFormat(), format);
}

// 测试自动筛选
TEST_F(WorksheetTest, AutoFilter) {
    worksheet->setAutoFilter(0, 0, 10, 5);
    
    EXPECT_TRUE(worksheet->hasAutoFilter());
    auto filter_range = worksheet->getAutoFilterRange();
    EXPECT_EQ(filter_range.first_row, 0);
    EXPECT_EQ(filter_range.first_col, 0);
    EXPECT_EQ(filter_range.last_row, 10);
    EXPECT_EQ(filter_range.last_col, 5);
    
    // 移除自动筛选
    worksheet->removeAutoFilter();
    EXPECT_FALSE(worksheet->hasAutoFilter());
}

// 测试冻结窗格
TEST_F(WorksheetTest, FreezePanes) {
    worksheet->freezePanes(1, 0);
    
    EXPECT_TRUE(worksheet->hasFrozenPanes());
    auto freeze_info = worksheet->getFreezeInfo();
    EXPECT_EQ(freeze_info.row, 1);
    EXPECT_EQ(freeze_info.col, 0);
    
    // 高级冻结窗格
    worksheet->freezePanes(2, 1, 2, 1);
    freeze_info = worksheet->getFreezeInfo();
    EXPECT_EQ(freeze_info.row, 2);
    EXPECT_EQ(freeze_info.col, 1);
    EXPECT_EQ(freeze_info.top_left_row, 2);
    EXPECT_EQ(freeze_info.top_left_col, 1);
}

// 测试分割窗格
TEST_F(WorksheetTest, SplitPanes) {
    worksheet->splitPanes(5, 2);
    
    EXPECT_TRUE(worksheet->hasFrozenPanes()); // 分割窗格也使用冻结窗格的数据结构
    auto freeze_info = worksheet->getFreezeInfo();
    EXPECT_EQ(freeze_info.row, 5);
    EXPECT_EQ(freeze_info.col, 2);
}

// 测试打印设置
TEST_F(WorksheetTest, PrintSettings) {
    // 设置打印区域
    worksheet->setPrintArea(0, 0, 20, 10);
    auto print_area = worksheet->getPrintArea();
    EXPECT_EQ(print_area.first_row, 0);
    EXPECT_EQ(print_area.first_col, 0);
    EXPECT_EQ(print_area.last_row, 20);
    EXPECT_EQ(print_area.last_col, 10);
    
    // 设置重复行
    worksheet->setRepeatRows(0, 2);
    auto repeat_rows = worksheet->getRepeatRows();
    EXPECT_EQ(repeat_rows.first, 0);
    EXPECT_EQ(repeat_rows.second, 2);
    
    // 设置重复列
    worksheet->setRepeatColumns(0, 1);
    auto repeat_cols = worksheet->getRepeatColumns();
    EXPECT_EQ(repeat_cols.first, 0);
    EXPECT_EQ(repeat_cols.second, 1);
    
    // 设置横向打印
    worksheet->setLandscape(true);
    EXPECT_TRUE(worksheet->isLandscape());
    
    // 设置页边距
    worksheet->setMargins(1.0, 1.0, 1.5, 1.5);
    auto margins = worksheet->getMargins();
    EXPECT_DOUBLE_EQ(margins.left, 1.0);
    EXPECT_DOUBLE_EQ(margins.right, 1.0);
    EXPECT_DOUBLE_EQ(margins.top, 1.5);
    EXPECT_DOUBLE_EQ(margins.bottom, 1.5);
    
    // 设置打印缩放
    worksheet->setPrintScale(80);
    EXPECT_EQ(worksheet->getPrintScale(), 80);
    
    // 设置适应页面
    worksheet->setFitToPages(1, 2);
    auto fit_pages = worksheet->getFitToPages();
    EXPECT_EQ(fit_pages.first, 1);
    EXPECT_EQ(fit_pages.second, 2);
    
    // 设置打印网格线
    worksheet->setPrintGridlines(true);
    EXPECT_TRUE(worksheet->isPrintGridlines());
    
    // 设置打印标题
    worksheet->setPrintHeadings(true);
    EXPECT_TRUE(worksheet->isPrintHeadings());
    
    // 设置页面居中
    worksheet->setCenterOnPage(true, true);
    EXPECT_TRUE(worksheet->isCenterHorizontally());
    EXPECT_TRUE(worksheet->isCenterVertically());
}

// 测试工作表保护
TEST_F(WorksheetTest, Protection) {
    EXPECT_FALSE(worksheet->isProtected());
    
    worksheet->protect("password123");
    EXPECT_TRUE(worksheet->isProtected());
    EXPECT_EQ(worksheet->getProtectionPassword(), "password123");
    
    worksheet->unprotect();
    EXPECT_FALSE(worksheet->isProtected());
}

// 测试视图设置
TEST_F(WorksheetTest, ViewSettings) {
    // 设置缩放
    worksheet->setZoom(150);
    EXPECT_EQ(worksheet->getZoom(), 150);
    
    // 显示/隐藏网格线
    worksheet->showGridlines(false);
    EXPECT_FALSE(worksheet->isGridlinesVisible());
    
    worksheet->showGridlines(true);
    EXPECT_TRUE(worksheet->isGridlinesVisible());
    
    // 显示/隐藏行列标题
    worksheet->showRowColHeaders(false);
    EXPECT_FALSE(worksheet->isRowColHeadersVisible());
    
    // 设置从右到左
    worksheet->setRightToLeft(true);
    EXPECT_TRUE(worksheet->isRightToLeft());
    
    // 设置选项卡选中
    worksheet->setTabSelected(true);
    EXPECT_TRUE(worksheet->isTabSelected());
    
    // 设置活动单元格
    worksheet->setActiveCell(5, 3);
    EXPECT_EQ(worksheet->getActiveCell(), "D6"); // 转换为Excel引用
    
    // 设置选择范围
    worksheet->setSelection(2, 1, 5, 4);
    EXPECT_EQ(worksheet->getSelection(), "B3:E6");
}

// 测试单元格检查
TEST_F(WorksheetTest, CellChecking) {
    EXPECT_FALSE(worksheet->hasCellAt(0, 0));
    
    worksheet->writeString(0, 0, "test");
    EXPECT_TRUE(worksheet->hasCellAt(0, 0));
    
    worksheet->writeString(5, 5, "");
    EXPECT_TRUE(worksheet->hasCellAt(5, 5)); // 空字符串也算有数据
}

// 测试清空操作
TEST_F(WorksheetTest, ClearOperations) {
    // 添加一些数据
    worksheet->writeString(0, 0, "A1");
    worksheet->writeString(0, 1, "B1");
    worksheet->writeString(1, 0, "A2");
    worksheet->writeString(1, 1, "B2");
    
    // 清空范围
    worksheet->clearRange(0, 0, 0, 1);
    EXPECT_FALSE(worksheet->hasCellAt(0, 0));
    EXPECT_FALSE(worksheet->hasCellAt(0, 1));
    EXPECT_TRUE(worksheet->hasCellAt(1, 0));
    EXPECT_TRUE(worksheet->hasCellAt(1, 1));
    
    // 清空整个工作表
    worksheet->clear();
    EXPECT_FALSE(worksheet->hasCellAt(1, 0));
    EXPECT_FALSE(worksheet->hasCellAt(1, 1));
    EXPECT_EQ(worksheet->getUsedRange().first, -1);
}

// 测试行列插入删除
TEST_F(WorksheetTest, InsertDeleteRowsColumns) {
    // 添加测试数据
    worksheet->writeString(0, 0, "A1");
    worksheet->writeString(1, 0, "A2");
    worksheet->writeString(2, 0, "A3");
    
    // 插入行
    worksheet->insertRows(1, 1);
    
    // 验证数据移动
    EXPECT_EQ(worksheet->getCell(0, 0).getStringValue(), "A1");
    EXPECT_TRUE(worksheet->getCell(1, 0).isEmpty()); // 插入的空行
    EXPECT_EQ(worksheet->getCell(2, 0).getStringValue(), "A2");
    EXPECT_EQ(worksheet->getCell(3, 0).getStringValue(), "A3");
    
    // 删除行
    worksheet->deleteRows(1, 1);
    
    // 验证数据恢复
    EXPECT_EQ(worksheet->getCell(0, 0).getStringValue(), "A1");
    EXPECT_EQ(worksheet->getCell(1, 0).getStringValue(), "A2");
    EXPECT_EQ(worksheet->getCell(2, 0).getStringValue(), "A3");
}

// 测试XML生成
TEST_F(WorksheetTest, XMLGeneration) {
    // 添加一些内容
    worksheet->writeString(0, 0, "Hello");
    worksheet->writeNumber(0, 1, 123.45);
    
    auto format = workbook->createFormat();
    format->setBold(true);
    worksheet->writeString(1, 0, "Bold Text", format);
    
    // 生成XML
    std::string xml = worksheet->generateXML();
    EXPECT_FALSE(xml.empty());
    EXPECT_NE(xml.find("<worksheet"), std::string::npos);
    EXPECT_NE(xml.find("<sheetData"), std::string::npos);
    EXPECT_NE(xml.find("Hello"), std::string::npos);
    
    // 生成关系XML
    worksheet->writeUrl(2, 0, "https://example.com", "Link");
    std::string rels_xml = worksheet->generateRelsXML();
    EXPECT_FALSE(rels_xml.empty());
    EXPECT_NE(rels_xml.find("<Relationships"), std::string::npos);
}

// 测试参数验证
TEST_F(WorksheetTest, ParameterValidation) {
    // 测试负数位置
    EXPECT_THROW(worksheet->writeString(-1, 0, "test"), std::invalid_argument);
    EXPECT_THROW(worksheet->writeString(0, -1, "test"), std::invalid_argument);
    
    // 测试超出Excel限制的位置
    EXPECT_THROW(worksheet->writeString(1048576, 0, "test"), std::invalid_argument);
    EXPECT_THROW(worksheet->writeString(0, 16384, "test"), std::invalid_argument);
    
    // 测试无效范围
    EXPECT_THROW(worksheet->mergeCells(5, 5, 2, 2), std::invalid_argument);
}

// 测试大量数据
TEST_F(WorksheetTest, LargeDataSet) {
    const int rows = 1000;
    const int cols = 10;
    
    // 写入大量数据
    for (int row = 0; row < rows; ++row) {
        for (int col = 0; col < cols; ++col) {
            worksheet->writeNumber(row, col, row * cols + col);
        }
    }
    
    // 验证使用范围
    auto [max_row, max_col] = worksheet->getUsedRange();
    EXPECT_EQ(max_row, rows - 1);
    EXPECT_EQ(max_col, cols - 1);
    
    // 验证部分数据
    EXPECT_DOUBLE_EQ(worksheet->getCell(0, 0).getNumberValue(), 0);
    EXPECT_DOUBLE_EQ(worksheet->getCell(100, 5).getNumberValue(), 100 * cols + 5);
    EXPECT_DOUBLE_EQ(worksheet->getCell(rows - 1, cols - 1).getNumberValue(), (rows - 1) * cols + (cols - 1));
}

// 测试性能（基本测试）
TEST_F(WorksheetTest, Performance) {
    auto start = std::chrono::high_resolution_clock::now();
    
    // 写入10000个单元格
    for (int i = 0; i < 10000; ++i) {
        int row = i / 100;
        int col = i % 100;
        worksheet->writeNumber(row, col, i);
    }
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);
    
    // 性能测试：10000个单元格应该在合理时间内完成（比如1秒）
    EXPECT_LT(duration.count(), 1000);
    
    // 验证数据正确性
    EXPECT_DOUBLE_EQ(worksheet->getCell(50, 50).getNumberValue(), 5050);
}