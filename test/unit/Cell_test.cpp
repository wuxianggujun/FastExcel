// FastExcel 库 - 高性能Excel文件处理库
// 组件：核心单元格处理模块测试
//
// 本文件实现了 Cell 类的单元测试，使用 Gtest 框架测试单元格的基本功能。

#include "fastexcel/core/Cell.hpp"
#include <gtest/gtest.h>
#include <string>

namespace fastexcel {
namespace core {

// Cell 类测试套件
class CellTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 测试前的设置
        cell = std::make_unique<Cell>();
    }

    void TearDown() override {
        // 测试后的清理
        cell.reset();
    }

    std::unique_ptr<Cell> cell;
};

// 测试默认构造函数
TEST_F(CellTest, DefaultConstructor) {
    EXPECT_EQ(cell->getType(), CellType::Empty);
    EXPECT_TRUE(cell->isEmpty());
    EXPECT_EQ(cell->getFormat(), nullptr);
}

// 测试字符串值设置和获取
TEST_F(CellTest, StringValue) {
    std::string testValue = "Hello, World!";
    cell->setValue(testValue);
    
    EXPECT_EQ(cell->getType(), CellType::String);
    EXPECT_TRUE(cell->isString());
    EXPECT_FALSE(cell->isEmpty());
    EXPECT_EQ(cell->getStringValue(), testValue);
}

// 测试数值设置和获取
TEST_F(CellTest, NumberValue) {
    double testValue = 42.5;
    cell->setValue(testValue);
    
    EXPECT_EQ(cell->getType(), CellType::Number);
    EXPECT_TRUE(cell->isNumber());
    EXPECT_FALSE(cell->isEmpty());
    EXPECT_DOUBLE_EQ(cell->getNumberValue(), testValue);
}

// 测试整数值设置和获取
TEST_F(CellTest, IntegerValue) {
    int testValue = 42;
    cell->setValue(testValue);
    
    EXPECT_EQ(cell->getType(), CellType::Number);
    EXPECT_TRUE(cell->isNumber());
    EXPECT_DOUBLE_EQ(cell->getNumberValue(), 42.0);
}

// 测试布尔值设置和获取
TEST_F(CellTest, BooleanValue) {
    cell->setValue(true);
    
    EXPECT_EQ(cell->getType(), CellType::Boolean);
    EXPECT_TRUE(cell->isBoolean());
    EXPECT_FALSE(cell->isEmpty());
    EXPECT_TRUE(cell->getBooleanValue());
    
    cell->setValue(false);
    EXPECT_FALSE(cell->getBooleanValue());
}

// 测试公式设置和获取
TEST_F(CellTest, Formula) {
    std::string testFormula = "=A1+B1";
    cell->setFormula(testFormula);
    
    EXPECT_EQ(cell->getType(), CellType::Formula);
    EXPECT_TRUE(cell->isFormula());
    EXPECT_FALSE(cell->isEmpty());
    EXPECT_EQ(cell->getFormula(), testFormula);
}

// 测试单元格清空
TEST_F(CellTest, Clear) {
    cell->setValue("Test");
    EXPECT_FALSE(cell->isEmpty());
    
    cell->clear();
    EXPECT_EQ(cell->getType(), CellType::Empty);
    EXPECT_TRUE(cell->isEmpty());
}

// 测试单元格复制构造函数
TEST_F(CellTest, CopyConstructor) {
    cell->setValue("Original");
    Cell copiedCell(*cell);
    
    EXPECT_EQ(copiedCell.getType(), cell->getType());
    EXPECT_EQ(copiedCell.getStringValue(), cell->getStringValue());
}

// 测试单元格赋值操作符
TEST_F(CellTest, AssignmentOperator) {
    cell->setValue("Original");
    Cell assignedCell;
    assignedCell = *cell;
    
    EXPECT_EQ(assignedCell.getType(), cell->getType());
    EXPECT_EQ(assignedCell.getStringValue(), cell->getStringValue());
}

// 测试单元格移动构造函数
TEST_F(CellTest, MoveConstructor) {
    cell->setValue("Original");
    Cell movedCell(std::move(*cell));
    
    EXPECT_EQ(movedCell.getType(), CellType::String);
    EXPECT_EQ(movedCell.getStringValue(), "Original");
}

// 测试单元格移动赋值操作符
TEST_F(CellTest, MoveAssignmentOperator) {
    cell->setValue("Original");
    Cell movedCell;
    movedCell = std::move(*cell);
    
    EXPECT_EQ(movedCell.getType(), CellType::String);
    EXPECT_EQ(movedCell.getStringValue(), "Original");
}

// 测试格式设置和获取
TEST_F(CellTest, Format) {
    auto format = std::make_shared<Format>();
    cell->setFormat(format);
    
    EXPECT_EQ(cell->getFormat(), format);
}

}} // namespace fastexcel::core