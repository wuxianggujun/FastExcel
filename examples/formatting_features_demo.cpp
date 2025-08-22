#include "fastexcel/utils/Logger.hpp"
/**
 * @file 04_formatting_example.cpp
 * @brief FastExcel格式化示例
 * 
 * 演示如何使用FastExcel进行单元格格式化
 */

#include "fastexcel/FastExcel.hpp"
#include "fastexcel/core/Path.hpp"
#include "fastexcel/core/StyleBuilder.hpp"
#include "fastexcel/core/FormatDescriptor.hpp"
#include <iostream>
#include <ctime>

int main() {
    try {
        // 初始化FastExcel库
        if (!fastexcel::initialize("logs/formatting_example.log", true)) {
            FASTEXCEL_LOG_ERROR("无法初始化FastExcel库");
            return -1;
        }
        
        // 创建工作簿
        auto workbook = fastexcel::core::Workbook::create(fastexcel::core::Path("formatting_example.xlsx"));
        if (!workbook) {
            FASTEXCEL_LOG_ERROR("无法创建工作簿");
            return -1;
        }
        
        // 添加工作表
        auto worksheet = workbook->addSheet("格式化示例");
        
        // ========== 使用新的样式系统创建各种格式 ==========
        
        // 标题格式：粗体、居中、蓝色背景（使用StyleBuilder）
        auto title_style = workbook->createStyleBuilder()
            .setBold(true)
            .setHorizontalAlign(fastexcel::HorizontalAlign::Center)
            .setVerticalAlign(fastexcel::VerticalAlign::Center)
            .setBackgroundColor(fastexcel::core::Color::fromRGB(0x4472C4))
            .setFontColor(fastexcel::core::Color::fromRGB(0xFFFFFF))
            .setFontSize(14)
            .build();
        int title_format_id = workbook->addStyle(title_style);
        
        // 表头格式：粗体、灰色背景、边框
        auto header_style = workbook->createStyleBuilder()
            .setBold(true)
            .setBackgroundColor(fastexcel::core::Color::fromRGB(0xD9D9D9))
            .setBorder(fastexcel::BorderStyle::Thin)
            .setHorizontalAlign(fastexcel::HorizontalAlign::Center)
            .build();
        int header_format_id = workbook->addStyle(header_style);
        
        // 数字格式：千分位分隔符
        auto number_style = workbook->createStyleBuilder()
            .setNumberFormat("#,##0.00")
            .setBorder(fastexcel::BorderStyle::Thin)
            .build();
        int number_format_id = workbook->addStyle(number_style);
        
        // 百分比格式
        auto percent_style = workbook->createStyleBuilder()
            .setNumberFormat("0.00%")
            .setBorder(fastexcel::BorderStyle::Thin)
            .build();
        int percent_format_id = workbook->addStyle(percent_style);
        
        // 日期格式
        auto date_style = workbook->createStyleBuilder()
            .setNumberFormat("yyyy-mm-dd")
            .setBorder(fastexcel::BorderStyle::Thin)
            .build();
        int date_format_id = workbook->addStyle(date_style);
        
        // 货币格式
        auto currency_style = workbook->createStyleBuilder()
            .setNumberFormat("¥#,##0.00")
            .setBorder(fastexcel::BorderStyle::Thin)
            .build();
        int currency_format_id = workbook->addStyle(currency_style);
        
        // 文本格式：左对齐、边框
        auto text_style = workbook->createStyleBuilder()
            .setHorizontalAlign(fastexcel::HorizontalAlign::Left)
            .setBorder(fastexcel::BorderStyle::Thin)
            .build();
        int text_format_id = workbook->addStyle(text_style);
        
        // 警告格式：红色背景、白色字体
        auto warning_style = workbook->createStyleBuilder()
            .setBackgroundColor(fastexcel::core::Color::fromRGB(0xFF0000))
            .setFontColor(fastexcel::core::Color::fromRGB(0xFFFFFF))
            .setBold(true)
            .setBorder(fastexcel::BorderStyle::Thin)
            .build();
        int warning_format_id = workbook->addStyle(warning_style);
        
        // ========== 写入数据并应用格式 ==========
        
        // 合并单元格作为标题
        worksheet->mergeCells(0, 0, 0, 5);
        worksheet->setValue(0, 0, std::string("销售数据报表"));
        worksheet->getCell(0, 0).setFormat(workbook->getStyle(title_format_id));
        worksheet->setRowHeight(0, 25);
        
        // 表头
        worksheet->setValue(2, 0, std::string("产品名称"));
        worksheet->getCell(2, 0).setFormat(workbook->getStyle(header_format_id));
        
        worksheet->setValue(2, 1, std::string("销售数量"));
        worksheet->getCell(2, 1).setFormat(workbook->getStyle(header_format_id));
        
        worksheet->setValue(2, 2, std::string("单价"));
        worksheet->getCell(2, 2).setFormat(workbook->getStyle(header_format_id));
        
        worksheet->setValue(2, 3, std::string("总金额"));
        worksheet->getCell(2, 3).setFormat(workbook->getStyle(header_format_id));
        
        worksheet->setValue(2, 4, std::string("增长率"));
        worksheet->getCell(2, 4).setFormat(workbook->getStyle(header_format_id));
        
        worksheet->setValue(2, 5, std::string("销售日期"));
        worksheet->getCell(2, 5).setFormat(workbook->getStyle(header_format_id));
        
        // 数据行
        worksheet->setValue(3, 0, std::string("笔记本电脑"));
        worksheet->getCell(3, 0).setFormat(workbook->getStyle(text_format_id));
        
        worksheet->setValue(3, 1, 150.0);
        worksheet->getCell(3, 1).setFormat(workbook->getStyle(number_format_id));
        
        worksheet->setValue(3, 2, 4999.99);
        worksheet->getCell(3, 2).setFormat(workbook->getStyle(currency_format_id));
        
        worksheet->getCell(3, 3).setFormula("B4*C4");
        worksheet->getCell(3, 3).setFormat(workbook->getStyle(currency_format_id));
        
        worksheet->setValue(3, 4, 0.15);
        worksheet->getCell(3, 4).setFormat(workbook->getStyle(percent_format_id));
        
        // 写入日期
        worksheet->setValue(3, 5, std::string("2024-01-15"));
        worksheet->getCell(3, 5).setFormat(workbook->getStyle(date_format_id));
        
        worksheet->setValue(4, 0, std::string("智能手机"));
        worksheet->getCell(4, 0).setFormat(workbook->getStyle(text_format_id));
        
        worksheet->setValue(4, 1, 300.0);
        worksheet->getCell(4, 1).setFormat(workbook->getStyle(number_format_id));
        
        worksheet->setValue(4, 2, 2999.00);
        worksheet->getCell(4, 2).setFormat(workbook->getStyle(currency_format_id));
        
        worksheet->getCell(4, 3).setFormula("B5*C5");
        worksheet->getCell(4, 3).setFormat(workbook->getStyle(currency_format_id));
        
        worksheet->setValue(4, 4, 0.25);
        worksheet->getCell(4, 4).setFormat(workbook->getStyle(percent_format_id));
        
        worksheet->setValue(4, 5, std::string("2024-01-20"));
        worksheet->getCell(4, 5).setFormat(workbook->getStyle(date_format_id));
        
        worksheet->setValue(5, 0, std::string("平板电脑"));
        worksheet->getCell(5, 0).setFormat(workbook->getStyle(text_format_id));
        
        worksheet->setValue(5, 1, 80.0);
        worksheet->getCell(5, 1).setFormat(workbook->getStyle(number_format_id));
        
        worksheet->setValue(5, 2, 1999.50);
        worksheet->getCell(5, 2).setFormat(workbook->getStyle(currency_format_id));
        
        worksheet->getCell(5, 3).setFormula("B6*C6");
        worksheet->getCell(5, 3).setFormat(workbook->getStyle(currency_format_id));
        
        worksheet->setValue(5, 4, -0.05); // 负增长用警告格式
        worksheet->getCell(5, 4).setFormat(workbook->getStyle(warning_format_id));
        
        worksheet->setValue(5, 5, std::string("2024-01-25"));
        worksheet->getCell(5, 5).setFormat(workbook->getStyle(date_format_id));
        
        // 总计行
        worksheet->setValue(7, 0, std::string("总计"));
        worksheet->getCell(7, 0).setFormat(workbook->getStyle(header_format_id));
        
        worksheet->getCell(7, 1).setFormula("SUM(B4:B6)");
        worksheet->getCell(7, 1).setFormat(workbook->getStyle(number_format_id));
        
        worksheet->setValue(7, 2, std::string(""));
        worksheet->getCell(7, 2).setFormat(workbook->getStyle(header_format_id));
        
        worksheet->getCell(7, 3).setFormula("SUM(D4:D6)");
        worksheet->getCell(7, 3).setFormat(workbook->getStyle(currency_format_id));
        
        worksheet->setValue(7, 4, std::string(""));
        worksheet->getCell(7, 4).setFormat(workbook->getStyle(header_format_id));
        
        worksheet->setValue(7, 5, std::string(""));
        worksheet->getCell(7, 5).setFormat(workbook->getStyle(header_format_id));
        
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
        
        // 设置文档属性（使用新API）
        workbook->setDocumentProperties(
            "销售数据报表",
            "格式化示例",
            "FastExcel",
            "FastExcel公司",
            "演示新的样式系统功能"
        );
        workbook->setKeywords("Excel, 格式化, 销售, 报表");
        
        // 添加自定义属性
        workbook->setProperty("部门", "销售部");
        workbook->setProperty("报表类型", "月度报表");
        workbook->setProperty("版本", 1.0);
        
        // 显示样式统计信息
        auto style_stats = workbook->getStyleStats();
        EXAMPLE_INFO("样式统计信息:");
        EXAMPLE_INFO("  - 总样式数: {}", style_stats.total_formats);
        EXAMPLE_INFO("  - 去重后样式数: {}", style_stats.unique_formats);
        EXAMPLE_INFO("  - 去重率: {:.1f}%", style_stats.deduplication_ratio * 100);
        
        // 保存文件
        if (!workbook->save()) {
            EXAMPLE_ERROR("保存文件失败");
            return -1;
        }
        
        EXAMPLE_INFO("格式化Excel文件创建成功: formatting_example.xlsx");
        
        // 清理资源
        fastexcel::cleanup();
        
    } catch (const std::exception& e) {
        EXAMPLE_ERROR("发生错误: {}", e.what());
        return -1;
    }
    
    return 0;
}
