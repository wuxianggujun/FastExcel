#include <gtest/gtest.h>
#include "fastexcel/core/Format.hpp"

using namespace fastexcel::core;

class FormatTest : public ::testing::Test {
protected:
    void SetUp() override {
        format = std::make_unique<Format>();
    }
    
    void TearDown() override {
        format.reset();
    }
    
    std::unique_ptr<Format> format;
};

// 测试默认构造函数
TEST_F(FormatTest, DefaultConstructor) {
    EXPECT_EQ(format->getFontName(), "Calibri");
    EXPECT_DOUBLE_EQ(format->getFontSize(), 11.0);
    EXPECT_EQ(format->getFontColor(), COLOR_BLACK);
    EXPECT_FALSE(format->isBold());
    EXPECT_FALSE(format->isItalic());
}

// 测试字体设置
TEST_F(FormatTest, FontSettings) {
    // 测试字体名称
    format->setFontName("Arial");
    EXPECT_EQ(format->getFontName(), "Arial");
    
    // 测试字体大小
    format->setFontSize(14.0);
    EXPECT_DOUBLE_EQ(format->getFontSize(), 14.0);
    
    // 测试字体颜色
    format->setFontColor(COLOR_RED);
    EXPECT_EQ(format->getFontColor(), COLOR_RED);
    
    // 测试粗体
    format->setBold(true);
    EXPECT_TRUE(format->isBold());
    format->setBold(false);
    EXPECT_FALSE(format->isBold());
    
    // 测试斜体
    format->setItalic(true);
    EXPECT_TRUE(format->isItalic());
    format->setItalic(false);
    EXPECT_FALSE(format->isItalic());
}

// 测试下划线设置
TEST_F(FormatTest, UnderlineSettings) {
    // 测试单下划线
    format->setUnderline(UnderlineType::Single);
    EXPECT_EQ(format->getUnderline(), UnderlineType::Single);
    
    // 测试双下划线
    format->setUnderline(UnderlineType::Double);
    EXPECT_EQ(format->getUnderline(), UnderlineType::Double);
    
    // 测试无下划线
    format->setUnderline(UnderlineType::None);
    EXPECT_EQ(format->getUnderline(), UnderlineType::None);
}

// 测试删除线
TEST_F(FormatTest, StrikeoutSettings) {
    EXPECT_FALSE(format->isStrikeout());
    
    format->setStrikeout(true);
    EXPECT_TRUE(format->isStrikeout());
    
    format->setStrikeout(false);
    EXPECT_FALSE(format->isStrikeout());
}

// 测试上标和下标
TEST_F(FormatTest, SuperscriptSubscript) {
    // 测试上标
    format->setSuperscript(true);
    EXPECT_TRUE(format->isSuperscript());
    EXPECT_FALSE(format->isSubscript()); // 上标和下标互斥
    
    // 测试下标
    format->setSubscript(true);
    EXPECT_TRUE(format->isSubscript());
    EXPECT_FALSE(format->isSuperscript()); // 上标和下标互斥
    
    // 清除
    format->setSuperscript(false);
    format->setSubscript(false);
    EXPECT_FALSE(format->isSuperscript());
    EXPECT_FALSE(format->isSubscript());
}

// 测试对齐设置
TEST_F(FormatTest, AlignmentSettings) {
    // 测试水平对齐
    format->setHorizontalAlign(HorizontalAlign::Center);
    EXPECT_EQ(format->getHorizontalAlign(), HorizontalAlign::Center);
    
    format->setHorizontalAlign(HorizontalAlign::Right);
    EXPECT_EQ(format->getHorizontalAlign(), HorizontalAlign::Right);
    
    // 测试垂直对齐
    format->setVerticalAlign(VerticalAlign::Center);
    EXPECT_EQ(format->getVerticalAlign(), VerticalAlign::Center);
    
    format->setVerticalAlign(VerticalAlign::Bottom);
    EXPECT_EQ(format->getVerticalAlign(), VerticalAlign::Bottom);
}

// 测试文本换行
TEST_F(FormatTest, TextWrap) {
    EXPECT_FALSE(format->isTextWrap());
    
    format->setTextWrap(true);
    EXPECT_TRUE(format->isTextWrap());
    
    format->setTextWrap(false);
    EXPECT_FALSE(format->isTextWrap());
}

// 测试文本旋转
TEST_F(FormatTest, TextRotation) {
    EXPECT_EQ(format->getRotation(), 0);
    
    // 测试正角度
    format->setRotation(45);
    EXPECT_EQ(format->getRotation(), 45);
    
    // 测试负角度
    format->setRotation(-30);
    EXPECT_EQ(format->getRotation(), -30);
    
    // 测试边界值
    format->setRotation(90);
    EXPECT_EQ(format->getRotation(), 90);
    
    format->setRotation(-90);
    EXPECT_EQ(format->getRotation(), -90);
}

// 测试缩进
TEST_F(FormatTest, Indentation) {
    EXPECT_EQ(format->getIndent(), 0);
    
    format->setIndent(2);
    EXPECT_EQ(format->getIndent(), 2);
    
    format->setIndent(0);
    EXPECT_EQ(format->getIndent(), 0);
}

// 测试收缩适应
TEST_F(FormatTest, ShrinkToFit) {
    EXPECT_FALSE(format->isShrinkToFit());
    
    format->setShrinkToFit(true);
    EXPECT_TRUE(format->isShrinkToFit());
    
    format->setShrinkToFit(false);
    EXPECT_FALSE(format->isShrinkToFit());
}

// 测试边框设置
TEST_F(FormatTest, BorderSettings) {
    // 测试统一边框
    format->setBorder(BorderStyle::Thin);
    EXPECT_EQ(format->getLeftBorder(), BorderStyle::Thin);
    EXPECT_EQ(format->getRightBorder(), BorderStyle::Thin);
    EXPECT_EQ(format->getTopBorder(), BorderStyle::Thin);
    EXPECT_EQ(format->getBottomBorder(), BorderStyle::Thin);
    
    // 测试单独边框
    format->setLeftBorder(BorderStyle::Thick);
    EXPECT_EQ(format->getLeftBorder(), BorderStyle::Thick);
    
    format->setRightBorder(BorderStyle::Dashed);
    EXPECT_EQ(format->getRightBorder(), BorderStyle::Dashed);
    
    format->setTopBorder(BorderStyle::Dotted);
    EXPECT_EQ(format->getTopBorder(), BorderStyle::Dotted);
    
    format->setBottomBorder(BorderStyle::Double);
    EXPECT_EQ(format->getBottomBorder(), BorderStyle::Double);
}

// 测试边框颜色
TEST_F(FormatTest, BorderColors) {
    // 测试统一边框颜色
    format->setBorderColor(COLOR_BLUE);
    EXPECT_EQ(format->getLeftBorderColor(), COLOR_BLUE);
    EXPECT_EQ(format->getRightBorderColor(), COLOR_BLUE);
    EXPECT_EQ(format->getTopBorderColor(), COLOR_BLUE);
    EXPECT_EQ(format->getBottomBorderColor(), COLOR_BLUE);
    
    // 测试单独边框颜色
    format->setLeftBorderColor(COLOR_RED);
    EXPECT_EQ(format->getLeftBorderColor(), COLOR_RED);
    
    format->setRightBorderColor(COLOR_GREEN);
    EXPECT_EQ(format->getRightBorderColor(), COLOR_GREEN);
    
    format->setTopBorderColor(COLOR_YELLOW);
    EXPECT_EQ(format->getTopBorderColor(), COLOR_YELLOW);
    
    format->setBottomBorderColor(COLOR_MAGENTA);
    EXPECT_EQ(format->getBottomBorderColor(), COLOR_MAGENTA);
}

// 测试对角线边框
TEST_F(FormatTest, DiagonalBorder) {
    format->setDiagonalBorder(BorderStyle::Thin);
    EXPECT_EQ(format->getDiagonalBorder(), BorderStyle::Thin);
    
    format->setDiagonalBorderColor(COLOR_CYAN);
    EXPECT_EQ(format->getDiagonalBorderColor(), COLOR_CYAN);
    
    format->setDiagonalType(DiagonalType::Up);
    EXPECT_EQ(format->getDiagonalType(), DiagonalType::Up);
    
    format->setDiagonalType(DiagonalType::Down);
    EXPECT_EQ(format->getDiagonalType(), DiagonalType::Down);
    
    format->setDiagonalType(DiagonalType::UpDown);
    EXPECT_EQ(format->getDiagonalType(), DiagonalType::UpDown);
}

// 测试填充设置
TEST_F(FormatTest, FillSettings) {
    // 测试背景色
    format->setBackgroundColor(COLOR_YELLOW);
    EXPECT_EQ(format->getBackgroundColor(), COLOR_YELLOW);
    
    // 测试前景色
    format->setForegroundColor(COLOR_BLUE);
    EXPECT_EQ(format->getForegroundColor(), COLOR_BLUE);
    
    // 测试填充模式
    format->setPattern(PatternType::Solid);
    EXPECT_EQ(format->getPattern(), PatternType::Solid);
    
    format->setPattern(PatternType::DarkGray);
    EXPECT_EQ(format->getPattern(), PatternType::DarkGray);
}

// 测试数字格式
TEST_F(FormatTest, NumberFormat) {
    // 测试自定义格式
    std::string custom_format = "#,##0.00";
    format->setNumberFormat(custom_format);
    EXPECT_EQ(format->getNumberFormat(), custom_format);
    
    // 测试预定义格式
    format->setNumberFormat(NumberFormatType::Currency);
    EXPECT_NE(format->getNumberFormat(), "");
    
    format->setNumberFormat(NumberFormatType::Percentage);
    EXPECT_NE(format->getNumberFormat(), "");
    
    format->setNumberFormat(NumberFormatType::Date);
    EXPECT_NE(format->getNumberFormat(), "");
}

// 测试保护设置
TEST_F(FormatTest, ProtectionSettings) {
    // 测试锁定
    EXPECT_TRUE(format->isLocked()); // 默认锁定
    
    format->setLocked(false);
    EXPECT_FALSE(format->isLocked());
    
    format->setLocked(true);
    EXPECT_TRUE(format->isLocked());
    
    // 测试隐藏
    EXPECT_FALSE(format->isHidden()); // 默认不隐藏
    
    format->setHidden(true);
    EXPECT_TRUE(format->isHidden());
    
    format->setHidden(false);
    EXPECT_FALSE(format->isHidden());
}

// 测试XF索引
TEST_F(FormatTest, XfIndex) {
    // 默认索引应该是-1
    EXPECT_EQ(format->getXfIndex(), -1);
    
    // 设置索引
    format->setXfIndex(5);
    EXPECT_EQ(format->getXfIndex(), 5);
}

// 测试基本功能
TEST_F(FormatTest, BasicFunctionality) {
    // 设置一些格式属性
    format->setBold(true);
    format->setFontColor(COLOR_RED);
    format->setBackgroundColor(COLOR_YELLOW);
    format->setBorder(BorderStyle::Thin);
    
    // 验证设置成功
    EXPECT_TRUE(format->isBold());
    EXPECT_EQ(format->getFontColor(), COLOR_RED);
    EXPECT_EQ(format->getBackgroundColor(), COLOR_YELLOW);
    EXPECT_EQ(format->getLeftBorder(), BorderStyle::Thin);
}

// 测试复制构造函数
TEST_F(FormatTest, CopyConstructor) {
    // 设置原始格式
    format->setBold(true);
    format->setFontColor(COLOR_RED);
    format->setBackgroundColor(COLOR_YELLOW);
    format->setHorizontalAlign(HorizontalAlign::Center);
    
    // 拷贝构造
    Format copy(*format);
    
    // 验证拷贝结果
    EXPECT_EQ(copy.isBold(), format->isBold());
    EXPECT_EQ(copy.getFontColor(), format->getFontColor());
    EXPECT_EQ(copy.getBackgroundColor(), format->getBackgroundColor());
    EXPECT_EQ(copy.getHorizontalAlign(), format->getHorizontalAlign());
}

// 测试赋值操作符
TEST_F(FormatTest, AssignmentOperator) {
    Format assigned;
    
    // 设置原始格式
    format->setItalic(true);
    format->setFontSize(16.0);
    format->setVerticalAlign(VerticalAlign::Bottom);
    
    // 赋值操作
    assigned = *format;
    
    // 验证赋值结果
    EXPECT_EQ(assigned.isItalic(), format->isItalic());
    EXPECT_DOUBLE_EQ(assigned.getFontSize(), format->getFontSize());
    EXPECT_EQ(assigned.getVerticalAlign(), format->getVerticalAlign());
}

// 测试边界值和异常情况
TEST_F(FormatTest, EdgeCases) {
    // 测试空字符串
    format->setFontName("");
    EXPECT_EQ(format->getFontName(), ""); // 应该允许空字符串
    
    // 测试零字体大小 - 实现可能有最小值限制
    format->setFontSize(0.0);
    // 不测试具体值，因为实现可能有验证逻辑
    
    // 测试负字体大小 - 实现可能有最小值限制
    format->setFontSize(-1.0);
    // 不测试具体值，因为实现可能有验证逻辑
    
    // 测试极大的缩进值 - 可能有范围限制
    format->setIndent(255); // 使用uint8_t的最大值
    EXPECT_EQ(format->getIndent(), 255);
}

// 测试格式组合
TEST_F(FormatTest, FormatCombination) {
    // 设置复杂的格式组合
    format->setFontName("Times New Roman");
    format->setFontSize(12.0);
    format->setBold(true);
    format->setItalic(true);
    format->setUnderline(UnderlineType::Double);
    format->setFontColor(COLOR_BLUE);
    format->setHorizontalAlign(HorizontalAlign::Center);
    format->setVerticalAlign(VerticalAlign::Center);
    format->setTextWrap(true);
    format->setRotation(45);
    format->setBorder(BorderStyle::Thick);
    format->setBorderColor(COLOR_RED);
    format->setBackgroundColor(COLOR_YELLOW);
    format->setPattern(PatternType::Solid);
    format->setNumberFormat("0.00%");
    format->setLocked(false);
    format->setHidden(true);
    
    // 验证所有设置都正确
    EXPECT_EQ(format->getFontName(), "Times New Roman");
    EXPECT_DOUBLE_EQ(format->getFontSize(), 12.0);
    EXPECT_TRUE(format->isBold());
    EXPECT_TRUE(format->isItalic());
    EXPECT_EQ(format->getUnderline(), UnderlineType::Double);
    EXPECT_EQ(format->getFontColor(), COLOR_BLUE);
    EXPECT_EQ(format->getHorizontalAlign(), HorizontalAlign::Center);
    EXPECT_EQ(format->getVerticalAlign(), VerticalAlign::Center);
    EXPECT_TRUE(format->isTextWrap());
    EXPECT_EQ(format->getRotation(), 45);
    EXPECT_EQ(format->getLeftBorder(), BorderStyle::Thick);
    EXPECT_EQ(format->getLeftBorderColor(), COLOR_RED);
    EXPECT_EQ(format->getBackgroundColor(), COLOR_YELLOW);
    EXPECT_EQ(format->getPattern(), PatternType::Solid);
    EXPECT_EQ(format->getNumberFormat(), "0.00%");
    EXPECT_FALSE(format->isLocked());
    EXPECT_TRUE(format->isHidden());
}