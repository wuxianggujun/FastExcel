/**
 * @file 04_formatting_example.cpp
 * @brief FastExcel格式化示例
 * 
 * 演示如何使用FastExcel进行单元格格式化
 */

#include "fastexcel/FastExcel.hpp"
#include <iostream>
#include <ctime>

int main() {
    try {
        // 初始化FastExcel库
        fastexcel::initialize();
        
        // 创建工作簿
        auto workbook = fastexcel::core::Workbook::create("formatting_example.xlsx");
        if (!workbook) {
            std::cerr << "无法创建工作簿" << std::endl;
            return -1;
        }
        
        // 添加工作表
        auto worksheet = workbook->addWorksheet("格式化示例");
        
        // ========== 创建各种格式 ==========
        
        // 标题格式：粗体、居中、蓝色背景
        auto title_format = workbook->createFormat();
        title_format->setBold(true);
        title_format->setHorizontalAlign(fastexcel::core::HorizontalAlign::Center);
        title_format->setVerticalAlign(fastexcel::core::VerticalAlign::Center);
        title_format->setBackgroundColor(fastexcel::core::COLOR_BLUE);
        title_format->setFontColor(fastexcel::core::COLOR_WHITE);
        title_format->setFontSize(14);
        
        // 表头格式：粗体、灰色背景、边框
        auto header_format = workbook->createFormat();
        header_format->setBold(true);
        header_format->setBackgroundColor(fastexcel::core::COLOR_GRAY);
        header_format->setBorder(fastexcel::core::BorderStyle::Thin);
        header_format->setHorizontalAlign(fastexcel::core::HorizontalAlign::Center);
        
        // 数字格式：千分位分隔符
        auto number_format = workbook->createFormat();
        number_format->setNumberFormat("#,##0.00");
        number_format->setBorder(fastexcel::core::BorderStyle::Thin);
        
        // 百分比格式
        auto percent_format = workbook->createFormat();
        percent_format->setNumberFormat("0.00%");
        percent_format->setBorder(fastexcel::core::BorderStyle::Thin);
        
        // 日期格式
        auto date_format = workbook->createFormat();
        date_format->setNumberFormat("yyyy-mm-dd");
        date_format->setBorder(fastexcel::core::BorderStyle::Thin);
        
        // 货币格式
        auto currency_format = workbook->createFormat();
        currency_format->setNumberFormat("¥#,##0.00");
        currency_format->setBorder(fastexcel::core::BorderStyle::Thin);
        
        // 文本格式：左对齐、边框
        auto text_format = workbook->createFormat();
        text_format->setHorizontalAlign(fastexcel::core::HorizontalAlign::Left);
        text_format->setBorder(fastexcel::core::BorderStyle::Thin);
        
        // 警告格式：红色背景、白色字体
        auto warning_format = workbook->createFormat();
        warning_format->setBackgroundColor(fastexcel::core::COLOR_RED);
        warning_format->setFontColor(fastexcel::core::COLOR_WHITE);
        warning_format->setBold(true);
        warning_format->setBorder(fastexcel::core::BorderStyle::Thin);
        
        // ========== 写入数据并应用格式 ==========
        
        // 合并单元格作为标题
        worksheet->mergeRange(0, 0, 0, 5, "销售数据报表", title_format);
        worksheet->setRowHeight(0, 25);
        
        // 表头
        worksheet->writeString(2, 0, "产品名称", header_format);
        worksheet->writeString(2, 1, "销售数量", header_format);
        worksheet->writeString(2, 2, "单价", header_format);
        worksheet->writeString(2, 3, "总金额", header_format);
        worksheet->writeString(2, 4, "增长率", header_format);
        worksheet->writeString(2, 5, "销售日期", header_format);
        
        // 数据行
        worksheet->writeString(3, 0, "笔记本电脑", text_format);
        worksheet->writeNumber(3, 1, 150, number_format);
        worksheet->writeNumber(3, 2, 4999.99, currency_format);
        worksheet->writeFormula(3, 3, "B4*C4", currency_format);
        worksheet->writeNumber(3, 4, 0.15, percent_format);
        
        // 写入日期
        std::tm date = {};
        date.tm_year = 124; // 2024年
        date.tm_mon = 0;    // 1月
        date.tm_mday = 15;  // 15日
        worksheet->writeDateTime(3, 5, date, date_format);
        
        worksheet->writeString(4, 0, "智能手机", text_format);
        worksheet->writeNumber(4, 1, 300, number_format);
        worksheet->writeNumber(4, 2, 2999.00, currency_format);
        worksheet->writeFormula(4, 3, "B5*C5", currency_format);
        worksheet->writeNumber(4, 4, 0.25, percent_format);
        
        date.tm_mday = 20;
        worksheet->writeDateTime(4, 5, date, date_format);
        
        worksheet->writeString(5, 0, "平板电脑", text_format);
        worksheet->writeNumber(5, 1, 80, number_format);
        worksheet->writeNumber(5, 2, 1999.50, currency_format);
        worksheet->writeFormula(5, 3, "B6*C6", currency_format);
        worksheet->writeNumber(5, 4, -0.05, warning_format); // 负增长用警告格式
        
        date.tm_mday = 25;
        worksheet->writeDateTime(5, 5, date, date_format);
        
        // 总计行
        worksheet->writeString(7, 0, "总计", header_format);
        worksheet->writeFormula(7, 1, "SUM(B4:B6)", number_format);
        worksheet->writeString(7, 2, "", header_format);
        worksheet->writeFormula(7, 3, "SUM(D4:D6)", currency_format);
        worksheet->writeString(7, 4, "", header_format);
        worksheet->writeString(7, 5, "", header_format);
        
        // ========== 设置列宽 ==========
        worksheet->setColumnWidth(0, 15);  // 产品名称
        worksheet->setColumnWidth(1, 12);  // 销售数量
        worksheet->setColumnWidth(2, 12);  // 单价
        worksheet->setColumnWidth(3, 15);  // 总金额
        worksheet->setColumnWidth(4, 10);  // 增长率
        worksheet->setColumnWidth(5, 12);  // 销售日期
        
        // ========== 设置打印选项 ==========
        worksheet->setPrintGridlines(true);
        worksheet->setPrintHeadings(true);
        worksheet->setLandscape(true);
        worksheet->setMargins(0.5, 0.5, 0.75, 0.75);
        
        // ========== 冻结窗格 ==========
        worksheet->freezePanes(3, 1); // 冻结表头
        
        // ========== 自动筛选 ==========
        worksheet->setAutoFilter(2, 0, 5, 5);
        
        // 设置文档属性
        workbook->setTitle("销售数据报表");
        workbook->setAuthor("FastExcel");
        workbook->setSubject("格式化示例");
        workbook->setKeywords("Excel, 格式化, 销售, 报表");
        
        // 添加自定义属性
        workbook->setCustomProperty("部门", "销售部");
        workbook->setCustomProperty("报表类型", "月度报表");
        workbook->setCustomProperty("版本", 1.0);
        
        // 保存文件
        if (!workbook->save()) {
            std::cerr << "保存文件失败" << std::endl;
            return -1;
        }
        
        std::cout << "格式化Excel文件创建成功: formatting_example.xlsx" << std::endl;
        
        // 清理资源
        fastexcel::cleanup();
        
    } catch (const std::exception& e) {
        std::cerr << "发生错误: " << e.what() << std::endl;
        return -1;
    }
    
    return 0;
}