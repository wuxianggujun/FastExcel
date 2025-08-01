#include <gtest/gtest.h>
#include "fastexcel/core/Cell.hpp"
#include "fastexcel/core/Format.hpp"
#include <memory>

using namespace fastexcel::core;

class CellTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 测试前的设置
    }
    
    void TearDown() override {
        // 测试后的清理
    }
};

// 测试默认构造函数
TEST_F(CellTest, DefaultConstructor) {
    Cell cell;
    
    EXPECT_EQ(cell.getType(), CellType::Empty);
    EXPECT_TRUE(cell.isEmpty());
    EXPECT_FALSE(cell.isString());
    EXPECT_FALSE(cell.isNumber());
    EXPECT_FALSE(cell.isBoolean());
    EXPECT_FALSE(cell.isFormula());
    EXPECT_FALSE(cell.hasHyperlink());
}

// 测试字符串值设置和获取
TEST_F(CellTest, StringValue) {
    Cell cell;
    std::string test_value = "Hello, World!";
    
    cell.setValue(test_value);
    
    EXPECT_EQ(cell.getType(), CellType::String);
    EXPECT_TRUE(cell.isString());
    EXPECT_FALSE(cell.isEmpty());
    EXPECT_EQ(cell.getStringValue(), test_value);
}

// 测试数字值设置和获取
TEST_F(CellTest, NumberValue) {
    Cell cell;
    double test_value = 123.456;
    
    cell.setValue(test_value);
    
    EXPECT_EQ(cell.getType(), CellType::Number);
    EXPECT_TRUE(cell.isNumber());
    EXPECT_FALSE(cell.isEmpty());
    EXPECT_DOUBLE_EQ(cell.getNumberValue(), test_value);
}

// 测试整数值设置
TEST_F(CellTest, IntegerValue) {
    Cell cell;
    int test_value = 42;
    
    cell.setValue(test_value);
    
    EXPECT_EQ(cell.getType(), CellType::Number);
    EXPECT_TRUE(cell.isNumber());
    EXPECT_DOUBLE_EQ(cell.getNumberValue(), static_cast<double>(test_value));
}

// 测试布尔值设置和获取
TEST_F(CellTest, BooleanValue) {
    Cell cell;
    
    // 测试 true
    cell.setValue(true);
    EXPECT_EQ(cell.getType(), CellType::Boolean);
    EXPECT_TRUE(cell.isBoolean());
    EXPECT_TRUE(cell.getBooleanValue());
    
    // 测试 false
    cell.setValue(false);
    EXPECT_EQ(cell.getType(), CellType::Boolean);
    EXPECT_TRUE(cell.isBoolean());
    EXPECT_FALSE(cell.getBooleanValue());
}

// 测试公式设置和获取
TEST_F(CellTest, FormulaValue) {
    Cell cell;
    std::string test_formula = "SUM(A1:A10)";
    
    cell.setFormula(test_formula);
    
    EXPECT_EQ(cell.getType(), CellType::Formula);
    EXPECT_TRUE(cell.isFormula());
    EXPECT_FALSE(cell.isEmpty());
    EXPECT_EQ(cell.getFormula(), test_formula);
}

// 测试超链接功能
TEST_F(CellTest, Hyperlink) {
    Cell cell;
    std::string test_url = "https://www.example.com";
    
    // 初始状态没有超链接
    EXPECT_FALSE(cell.hasHyperlink());
    EXPECT_TRUE(cell.getHyperlink().empty());
    
    // 设置超链接
    cell.setHyperlink(test_url);
    EXPECT_TRUE(cell.hasHyperlink());
    EXPECT_EQ(cell.getHyperlink(), test_url);
}

// 测试格式设置和获取
TEST_F(CellTest, Format) {
    Cell cell;
    auto format = std::make_shared<Format>();
    
    // 初始状态没有格式
    EXPECT_EQ(cell.getFormat(), nullptr);
    
    // 设置格式
    cell.setFormat(format);
    EXPECT_EQ(cell.getFormat(), format);
}

// 测试清空单元格
TEST_F(CellTest, Clear) {
    Cell cell;
    auto format = std::make_shared<Format>();
    
    // 设置各种数据
    cell.setValue("test");
    cell.setFormat(format);
    cell.setHyperlink("https://example.com");
    
    // 验证数据已设置
    EXPECT_FALSE(cell.isEmpty());
    EXPECT_TRUE(cell.hasHyperlink());
    EXPECT_NE(cell.getFormat(), nullptr);
    
    // 清空单元格
    cell.clear();
    
    // 验证已清空
    EXPECT_TRUE(cell.isEmpty());
    EXPECT_EQ(cell.getType(), CellType::Empty);
    EXPECT_FALSE(cell.hasHyperlink());
    EXPECT_EQ(cell.getFormat(), nullptr);
}

// 测试拷贝构造函数
TEST_F(CellTest, CopyConstructor) {
    Cell original;
    auto format = std::make_shared<Format>();
    
    // 设置原始单元格
    original.setValue("test value");
    original.setFormat(format);
    original.setHyperlink("https://example.com");
    
    // 拷贝构造
    Cell copy(original);
    
    // 验证拷贝结果
    EXPECT_EQ(copy.getType(), original.getType());
    EXPECT_EQ(copy.getStringValue(), original.getStringValue());
    EXPECT_EQ(copy.getFormat(), original.getFormat());
    EXPECT_EQ(copy.getHyperlink(), original.getHyperlink());
}

// 测试赋值操作符
TEST_F(CellTest, AssignmentOperator) {
    Cell original;
    Cell assigned;
    auto format = std::make_shared<Format>();
    
    // 设置原始单元格
    original.setValue(42.0);
    original.setFormat(format);
    original.setHyperlink("https://example.com");
    
    // 赋值操作
    assigned = original;
    
    // 验证赋值结果
    EXPECT_EQ(assigned.getType(), original.getType());
    EXPECT_DOUBLE_EQ(assigned.getNumberValue(), original.getNumberValue());
    EXPECT_EQ(assigned.getFormat(), original.getFormat());
    EXPECT_EQ(assigned.getHyperlink(), original.getHyperlink());
}

// 测试自赋值
TEST_F(CellTest, SelfAssignment) {
    Cell cell;
    auto format = std::make_shared<Format>();
    
    cell.setValue("test");
    cell.setFormat(format);
    
    // 自赋值
    cell = cell;
    
    // 验证数据未损坏
    EXPECT_EQ(cell.getStringValue(), "test");
    EXPECT_EQ(cell.getFormat(), format);
}

// 测试移动语义
TEST_F(CellTest, MoveSemantics) {
    Cell original;
    auto format = std::make_shared<Format>();
    
    original.setValue("test value");
    original.setFormat(format);
    original.setHyperlink("https://example.com");
    
    // 移动构造
    Cell moved(std::move(original));
    
    // 验证移动结果
    EXPECT_EQ(moved.getStringValue(), "test value");
    EXPECT_EQ(moved.getFormat(), format);
    EXPECT_EQ(moved.getHyperlink(), "https://example.com");
    
    // 注意：移动后原对象的状态是未定义的，不应该测试原对象
}

// 测试类型转换边界情况
TEST_F(CellTest, TypeConversionEdgeCases) {
    Cell cell;
    
    // 空单元格的各种获取方法应该返回默认值
    EXPECT_EQ(cell.getStringValue(), "");
    EXPECT_DOUBLE_EQ(cell.getNumberValue(), 0.0);
    EXPECT_FALSE(cell.getBooleanValue());
    EXPECT_EQ(cell.getFormula(), "");
    
    // 设置字符串后，其他类型应该返回默认值
    cell.setValue("hello");
    EXPECT_EQ(cell.getStringValue(), "hello");
    EXPECT_DOUBLE_EQ(cell.getNumberValue(), 0.0);
    EXPECT_FALSE(cell.getBooleanValue());
}

// 测试空字符串和空值
TEST_F(CellTest, EmptyStringAndValues) {
    Cell cell;
    
    // 设置空字符串
    cell.setValue("");
    EXPECT_EQ(cell.getType(), CellType::String);
    EXPECT_TRUE(cell.isString());
    EXPECT_EQ(cell.getStringValue(), "");
    
    // 设置零值
    cell.setValue(0.0);
    EXPECT_EQ(cell.getType(), CellType::Number);
    EXPECT_TRUE(cell.isNumber());
    EXPECT_DOUBLE_EQ(cell.getNumberValue(), 0.0);
}

// 测试特殊数值
TEST_F(CellTest, SpecialNumbers) {
    Cell cell;
    
    // 测试负数
    cell.setValue(-123.456);
    EXPECT_DOUBLE_EQ(cell.getNumberValue(), -123.456);
    
    // 测试很大的数
    cell.setValue(1e10);
    EXPECT_DOUBLE_EQ(cell.getNumberValue(), 1e10);
    
    // 测试很小的数
    cell.setValue(1e-10);
    EXPECT_DOUBLE_EQ(cell.getNumberValue(), 1e-10);
}