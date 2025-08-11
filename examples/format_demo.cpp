#include "fastexcel/FastExcel.hpp"
#include <iostream>

using namespace fastexcel;
using namespace fastexcel::core;

int main() {
    try {
        // 创建工作簿
        auto workbook = Workbook::create(Path("format_demo.xlsx"));
        if (!workbook) {
            std::cout << "Failed to create workbook" << std::endl;
            return 1;
        }
        
        auto worksheet = workbook->addWorksheet("格式演示");
        
        std::cout << "=== FastExcel 自动换行与数字格式演示 ===" << std::endl;
        
        // ========== 1. 自动换行演示 ==========
        std::cout << "\n📝 1. 自动换行功能:" << std::endl;
        
        auto wrapStyle = workbook->createStyleBuilder()
            .textWrap(true)                    // 🎯 自动换行
            .fontName("Arial")
            .fontSize(11)
            .verticalAlign(core::VerticalAlign::Top)
            .build();
        
        int wrapStyleId = workbook->addStyle(wrapStyle);
        
        worksheet->writeString(0, 0, "这是一段很长的文本\\n会自动换行显示\\n支持多行内容");
        auto& cell1 = worksheet->getCell(0, 0);
        cell1.setFormat(workbook->getStyles().getFormat(wrapStyleId));
        
        std::cout << "   ✅ 设置单元格自动换行: .textWrap(true)" << std::endl;
        
        // ========== 2. 数字格式演示 ==========
        std::cout << "\n💰 2. 数字格式功能:" << std::endl;
        
        // 2位小数格式
        auto decimalStyle = workbook->createStyleBuilder()
            .numberFormat("0.00")              // 🎯 自定义2位小数格式
            .rightAlign()
            .build();
        
        // 百分比格式
        auto percentStyle = workbook->createStyleBuilder()
            .percentage()                      // 🎯 内置百分比格式
            .rightAlign()
            .fontColor(Color(0, 128, 0))       // 绿色
            .build();
        
        // 货币格式
        auto currencyStyle = workbook->createStyleBuilder()
            .currency()                        // 🎯 内置货币格式
            .rightAlign()
            .fontColor(Color(0, 0, 255))       // 蓝色
            .build();
        
        // 科学计数法
        auto scientificStyle = workbook->createStyleBuilder()
            .scientific()                      // 🎯 内置科学计数法
            .rightAlign()
            .build();
        
        // 自定义千分位格式
        auto thousandStyle = workbook->createStyleBuilder()
            .numberFormat("#,##0")             // 🎯 千分位格式
            .rightAlign()
            .bold()
            .build();
        
        // 日期格式
        auto dateStyle = workbook->createStyleBuilder()
            .date()                           // 🎯 内置日期格式
            .centerAlign()
            .build();
        
        // 添加样式到工作簿
        int decimalStyleId = workbook->addStyle(decimalStyle);
        int percentStyleId = workbook->addStyle(percentStyle);
        int currencyStyleId = workbook->addStyle(currencyStyle);
        int scientificStyleId = workbook->addStyle(scientificStyle);
        int thousandStyleId = workbook->addStyle(thousandStyle);
        int dateStyleId = workbook->addStyle(dateStyle);
        
        // 写入标题行
        worksheet->writeString(2, 0, "数值类型");
        worksheet->writeString(2, 1, "原始值");
        worksheet->writeString(2, 2, "格式化后");
        worksheet->writeString(2, 3, "格式代码");
        
        int row = 3;
        
        // 2位小数演示
        worksheet->writeString(row, 0, "2位小数");
        worksheet->writeNumber(row, 1, 123.456789);
        worksheet->writeNumber(row, 2, 123.456789);
        worksheet->getCell(row, 2).setFormat(workbook->getStyles().getFormat(decimalStyleId));
        worksheet->writeString(row, 3, "0.00");
        std::cout << "   ✅ 2位小数格式: .numberFormat(\"0.00\")" << std::endl;
        row++;
        
        // 百分比演示
        worksheet->writeString(row, 0, "百分比");
        worksheet->writeNumber(row, 1, 0.85);
        worksheet->writeNumber(row, 2, 0.85);
        worksheet->getCell(row, 2).setFormat(workbook->getStyles().getFormat(percentStyleId));
        worksheet->writeString(row, 3, "0.00%");
        std::cout << "   ✅ 百分比格式: .percentage()" << std::endl;
        row++;
        
        // 货币演示
        worksheet->writeString(row, 0, "货币");
        worksheet->writeNumber(row, 1, 1234.56);
        worksheet->writeNumber(row, 2, 1234.56);
        worksheet->getCell(row, 2).setFormat(workbook->getStyles().getFormat(currencyStyleId));
        worksheet->writeString(row, 3, "¤#,##0.00");
        std::cout << "   ✅ 货币格式: .currency()" << std::endl;
        row++;
        
        // 科学计数法演示
        worksheet->writeString(row, 0, "科学计数法");
        worksheet->writeNumber(row, 1, 1234567.89);
        worksheet->writeNumber(row, 2, 1234567.89);
        worksheet->getCell(row, 2).setFormat(workbook->getStyles().getFormat(scientificStyleId));
        worksheet->writeString(row, 3, "0.00E+00");
        std::cout << "   ✅ 科学计数法: .scientific()" << std::endl;
        row++;
        
        // 千分位演示
        worksheet->writeString(row, 0, "千分位");
        worksheet->writeNumber(row, 1, 9876543);
        worksheet->writeNumber(row, 2, 9876543);
        worksheet->getCell(row, 2).setFormat(workbook->getStyles().getFormat(thousandStyleId));
        worksheet->writeString(row, 3, "#,##0");
        std::cout << "   ✅ 千分位格式: .numberFormat(\"#,##0\")" << std::endl;
        row++;
        
        // 更多自定义格式演示
        std::cout << "\n🎨 3. 更多自定义格式:" << std::endl;
        
        // 自定义格式: 正数绿色，负数红色
        auto customStyle = workbook->createStyleBuilder()
            .numberFormat("[GREEN]0.00;[RED]-0.00")  // 🎯 条件格式
            .rightAlign()
            .build();
        
        int customStyleId = workbook->addStyle(customStyle);
        
        worksheet->writeString(row, 0, "条件颜色");
        worksheet->writeNumber(row, 1, -456.78);
        worksheet->writeNumber(row, 2, -456.78);
        worksheet->getCell(row, 2).setFormat(workbook->getStyles().getFormat(customStyleId));
        worksheet->writeString(row, 3, "[GREEN]0.00;[RED]-0.00");
        std::cout << "   ✅ 条件格式: .numberFormat(\"[GREEN]0.00;[RED]-0.00\")" << std::endl;
        
        // 保存文件
        workbook->save();
        workbook->close();
        
        std::cout << "\n🎉 FastExcel 完全支持所有格式功能!" << std::endl;
        std::cout << "📋 可用的格式设置API:" << std::endl;
        std::cout << "   📝 自动换行: .textWrap(true)" << std::endl;
        std::cout << "   💰 货币格式: .currency()" << std::endl;
        std::cout << "   📊 百分比: .percentage()" << std::endl;
        std::cout << "   🔢 小数位: .numberFormat(\"0.00\")" << std::endl;
        std::cout << "   🔬 科学计数: .scientific()" << std::endl;
        std::cout << "   📅 日期格式: .date()" << std::endl;
        std::cout << "   🎨 自定义格式: .numberFormat(\"[GREEN]0.00;[RED]-0.00\")" << std::endl;
        
        std::cout << "\n✨ 生成的Excel文件: format_demo.xlsx" << std::endl;
        
    } catch (const std::exception& e) {
        std::cout << "❌ 错误: " << e.what() << std::endl;
        return 1;
    }
    
    return 0;
}