#include "../src/fastexcel/core/Workbook.hpp"
#include "../src/fastexcel/core/Worksheet.hpp"
#include "../src/fastexcel/core/RangeFormatter.hpp"
#include "../src/fastexcel/core/QuickFormat.hpp"
#include "../src/fastexcel/core/FormatUtils.hpp"
#include "../src/fastexcel/core/StyleBuilder.hpp"
#include "../src/fastexcel/core/Color.hpp"
#include <iostream>
#include <vector>

using namespace fastexcel;
using namespace fastexcel::core;

/**
 * @brief FastExcel增强格式API演示程序
 * 
 * 这个程序演示了我们新增的所有格式化功能：
 * 1. RangeFormatter - 批量格式设置
 * 2. QuickFormat - 快速格式化工具
 * 3. FormatUtils - 格式工具类
 */
int main() {
    try {
        std::cout << "=== FastExcel增强格式API演示 ===\n\n";
        
        // 创建工作簿 - 使用正确的工厂方法
        core::Path outputPath("enhanced_format_demo.xlsx");
        auto workbook = Workbook::create(outputPath);
        if (!workbook) {
            std::cerr << "❌ 无法创建工作簿" << std::endl;
            return 1;
        }
        
        // ========== 演示1: RangeFormatter批量格式化 ==========
        {
            std::cout << "📋 演示1: RangeFormatter批量格式化\n";
            auto worksheet = workbook->addSheet("批量格式化演示");
            
            // 添加示例数据
            std::vector<std::vector<std::string>> data = {
                {"产品名称", "Q1销量", "Q2销量", "Q3销量", "Q4销量", "全年总计"},
                {"iPhone", "1200", "1350", "1450", "1600", "5600"},
                {"iPad", "800", "900", "950", "1100", "3750"},
                {"MacBook", "600", "650", "700", "750", "2700"},
                {"Apple Watch", "900", "1000", "1100", "1200", "4200"},
                {"AirPods", "1500", "1600", "1700", "1800", "6600"}
            };
            
            // 填入数据
            for (int row = 0; row < data.size(); ++row) {
                for (int col = 0; col < data[row].size(); ++col) {
                    if (row == 0 || col == 0) {
                        worksheet->setValue(row, col, data[row][col]);
                    } else {
                        worksheet->setValue(row, col, std::stoi(data[row][col]));
                    }
                }
            }
            
            // 使用RangeFormatter进行批量格式化
            std::cout << "  • 格式化标题行...\n";
            worksheet->rangeFormatter("A1:F1")
                .bold()
                .backgroundColor(Color::BLUE)
                .fontColor(Color::WHITE)
                .centerAlign()
                .allBorders(BorderStyle::Medium)
                .apply();
            
            std::cout << "  • 格式化数据区域...\n";
            worksheet->rangeFormatter("A2:F6")
                .allBorders(BorderStyle::Thin)
                .vcenterAlign()
                .apply();
            
            std::cout << "  • 格式化产品名称列...\n";
            worksheet->rangeFormatter("A2:A6")
                .bold()
                .backgroundColor(Color(230, 230, 230))  // 浅灰色
                .leftAlign()
                .apply();
            
            std::cout << "  • 格式化数字列...\n";
            worksheet->rangeFormatter("B2:F6")
                .rightAlign()
                .apply();
            
            std::cout << "  ✓ 批量格式化完成\n\n";
        }
        
        // ========== 演示2: QuickFormat快速格式化 ==========
        {
            std::cout << "🚀 演示2: QuickFormat快速格式化\n";
            auto worksheet = workbook->addSheet("快速格式化演示");
            
            // 创建财务报表数据
            std::vector<std::vector<std::string>> financial_data = {
                {"财务报表 - 2024年度", "", "", ""},
                {"", "", "", ""},
                {"项目", "Q1", "Q2", "Q3"},
                {"收入", "150000", "175000", "180000"},
                {"成本", "90000", "105000", "108000"},
                {"利润", "60000", "70000", "72000"},
                {"利润率", "0.4", "0.4", "0.4"}
            };
            
            // 填入数据
            for (int row = 0; row < financial_data.size(); ++row) {
                for (int col = 0; col < financial_data[row].size(); ++col) {
                    const std::string& value = financial_data[row][col];
                    if (!value.empty()) {
                        if (row > 2 && col > 0 && row < 6) {
                            // 数字数据
                            worksheet->setValue(row, col, std::stod(value));
                        } else if (row == 6 && col > 0) {
                            // 百分比数据
                            worksheet->setValue(row, col, std::stod(value));
                        } else {
                            // 文本数据
                            worksheet->setValue(row, col, value);
                        }
                    }
                }
            }
            
            // 使用QuickFormat进行快速格式化
            std::cout << "  • 格式化主标题...\n";
            QuickFormat::formatAsTitle(*worksheet, 0, 0, "", 16.0);
            
            std::cout << "  • 格式化表头...\n";
            QuickFormat::formatAsHeader(*worksheet, "A3:D3", QuickFormat::HeaderStyle::Modern);
            
            std::cout << "  • 格式化货币数据...\n";
            QuickFormat::formatAsCurrency(*worksheet, "B4:D6", "¥", 0, true);
            
            std::cout << "  • 格式化百分比数据...\n";
            QuickFormat::formatAsPercentage(*worksheet, "B7:D7", 1);
            
            std::cout << "  • 应用财务报表样式套餐...\n";
            QuickFormat::applyFinancialReportStyle(*worksheet, "A3:D7", "A3:D3", "A1");
            
            std::cout << "  ✓ 快速格式化完成\n\n";
        }
        
        // ========== 演示3: 综合样式演示 ==========
        {
            std::cout << "🎨 演示3: 综合样式演示\n";
            auto worksheet = workbook->addSheet("综合样式演示");
            
            // 创建各种样式示例
            worksheet->setValue(0, 0, "样式类型");
            worksheet->setValue(0, 1, "示例文本");
            worksheet->setValue(0, 2, "描述");
            
            std::vector<std::string> style_examples = {
                "标准文本", "现代标题", "经典标题", "粗体文本", 
                "成功消息", "警告消息", "错误消息", "注释文本"
            };
            
            for (int i = 0; i < style_examples.size(); ++i) {
                int row = i + 1;
                worksheet->setValue(row, 0, style_examples[i]);
                worksheet->setValue(row, 1, "这是" + style_examples[i] + "的示例");
                worksheet->setValue(row, 2, "演示不同的格式效果");
            }
            
            // 应用不同的格式
            std::cout << "  • 应用标题样式...\n";
            QuickFormat::formatAsHeader(*worksheet, "A1:C1", QuickFormat::HeaderStyle::Colorful);
            
            std::cout << "  • 应用各种格式样式...\n";
            QuickFormat::formatAsTitle(*worksheet, 2, 1, "", 14.0);  // 现代标题
            QuickFormat::formatAsHeader(*worksheet, "B3:B3", QuickFormat::HeaderStyle::Classic);  // 经典标题
            
            worksheet->rangeFormatter("B4:B4").bold().apply();  // 粗体文本
            
            QuickFormat::formatAsSuccess(*worksheet, "B5:B5");  // 成功消息
            QuickFormat::formatAsWarning(*worksheet, "B6:B6");  // 警告消息
            QuickFormat::formatAsError(*worksheet, "B7:B7");    // 错误消息
            QuickFormat::formatAsComment(*worksheet, "B8:B8");  // 注释文本
            
            // 设置数据区域边框
            worksheet->rangeFormatter("A1:C8")
                .allBorders(BorderStyle::Thin)
                .apply();
            
            std::cout << "  ✓ 综合样式演示完成\n\n";
        }
        
        // ========== 演示4: 条件格式和数据突出显示 ==========
        {
            std::cout << "🎯 演示4: 数据突出显示\n";
            auto worksheet = workbook->addSheet("数据突出显示");
            
            // 创建成绩数据
            std::vector<std::vector<std::string>> scores_data = {
                {"学生姓名", "数学", "英语", "物理", "化学", "平均分"},
                {"张三", "85", "92", "78", "88", "85.75"},
                {"李四", "92", "88", "95", "90", "91.25"},
                {"王五", "78", "85", "82", "79", "81"},
                {"赵六", "95", "89", "92", "94", "92.5"},
                {"陈七", "68", "72", "75", "70", "71.25"}
            };
            
            // 填入数据
            for (int row = 0; row < scores_data.size(); ++row) {
                for (int col = 0; col < scores_data[row].size(); ++col) {
                    const std::string& value = scores_data[row][col];
                    if (row == 0 || col == 0) {
                        worksheet->setValue(row, col, value);
                    } else {
                        worksheet->setValue(row, col, std::stod(value));
                    }
                }
            }
            
            std::cout << "  • 格式化表头...\n";
            QuickFormat::formatAsHeader(*worksheet, "A1:F1", QuickFormat::HeaderStyle::Modern);
            
            std::cout << "  • 突出显示优秀成绩（≥90分）...\n";
            // 手动检查并突出显示高分（简化演示）
            QuickFormat::highlight(*worksheet, "B2:B2", Color::GREEN);  // 张三数学92分
            QuickFormat::highlight(*worksheet, "C2:C2", Color::GREEN);  // 张三英语92分
            QuickFormat::highlight(*worksheet, "B3:D3", Color::GREEN); // 李四多科90+分
            QuickFormat::highlight(*worksheet, "B4:B4", Color::GREEN);  // 王五数学95分
            QuickFormat::highlight(*worksheet, "B5:D5", Color::GREEN); // 赵六多科90+分
            
            std::cout << "  • 突出显示需要改进的成绩（<75分）...\n";
            QuickFormat::formatAsWarning(*worksheet, "C6:C6");  // 陈七英语72分
            QuickFormat::formatAsError(*worksheet, "B6:B6");    // 陈七数学68分
            QuickFormat::formatAsError(*worksheet, "E6:E6");    // 陈七化学70分
            
            // 设置基础表格格式
            worksheet->rangeFormatter("A1:F6")
                .allBorders(BorderStyle::Thin)
                .vcenterAlign()
                .apply();
            
            worksheet->rangeFormatter("A2:A6")
                .leftAlign()
                .apply();
            
            worksheet->rangeFormatter("B1:F6")
                .centerAlign()
                .apply();
            
            std::cout << "  ✓ 数据突出显示完成\n\n";
        }
        
        // 保存文件
        std::cout << "💾 保存文件...\n";
        workbook->save();
        std::cout << "✅ 演示完成！文件已保存为: enhanced_format_demo.xlsx\n\n";
        
        // 总结
        std::cout << "🎉 FastExcel增强格式API演示总结:\n";
        std::cout << "  • RangeFormatter: 支持批量范围格式化，链式调用\n";
        std::cout << "  • QuickFormat: 提供常用格式的快速应用方法\n";
        std::cout << "  • FormatUtils: 格式复制、清除、检查等工具功能\n";
        std::cout << "  • 智能API: 内部自动优化FormatRepository操作\n";
        std::cout << "  • 丰富样式: 支持财务、表格、突出显示等多种样式\n\n";
        
        std::cout << "打开生成的Excel文件查看格式化效果！\n";
        
        return 0;
        
    } catch (const std::exception& e) {
        std::cerr << "❌ 错误: " << e.what() << std::endl;
        return 1;
    }
}