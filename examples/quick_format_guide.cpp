#include "../src/fastexcel/core/Workbook.hpp"
#include "../src/fastexcel/core/Worksheet.hpp"
#include "../src/fastexcel/core/RangeFormatter.hpp"
#include "../src/fastexcel/core/QuickFormat.hpp"
#include "../src/fastexcel/core/StyleBuilder.hpp"
#include "../src/fastexcel/core/Color.hpp"
#include <iostream>

using namespace fastexcel;
using namespace fastexcel::core;

/**
 * @brief FastExcel格式化API快速上手指南
 * 
 * 这个简单的例子展示了新API的基本用法，
 * 帮助用户快速掌握新的格式化功能。
 */
int main() {
    try {
        std::cout << "FastExcel格式化API快速上手指南\n";
        std::cout << "==================================\n\n";
        
        // 1. 创建工作簿和工作表 - 使用正确的工厂方法
        core::Path outputPath("api_guide_demo.xlsx");
        auto workbook = Workbook::create(outputPath);
        if (!workbook) {
            std::cerr << "❌ 无法创建工作簿" << std::endl;
            return 1;
        }
        
        auto worksheet = workbook->addSheet("API使用指南");
        if (!worksheet) {
            std::cerr << "❌ 无法创建工作表" << std::endl;
            return 1;
        }
        
        // 2. 添加数据
        worksheet->setValue(0, 0, "FastExcel新API使用指南");
        worksheet->setValue(2, 0, "功能");
        worksheet->setValue(2, 1, "说明");
        worksheet->setValue(2, 2, "示例");
        
        worksheet->setValue(3, 0, "RangeFormatter");
        worksheet->setValue(3, 1, "批量格式化范围");
        worksheet->setValue(3, 2, "worksheet.rangeFormatter(\"A1:C3\")");
        
        worksheet->setValue(4, 0, "QuickFormat");
        worksheet->setValue(4, 1, "快速应用常用格式");
        worksheet->setValue(4, 2, "QuickFormat::formatAsCurrency()");
        
        worksheet->setValue(5, 0, "智能API");
        worksheet->setValue(5, 1, "自动优化性能");
        worksheet->setValue(5, 2, "内部自动处理FormatRepository");
        
        // 3. 使用新API进行格式化
        std::cout << "使用新API进行格式化...\n";
        
        // 3.1 格式化主标题 - 演示QuickFormat
        QuickFormat::formatAsTitle(*worksheet, 0, 0, "", 18.0);
        std::cout << "✓ 主标题格式化完成\n";
        
        // 3.2 格式化表头 - 演示RangeFormatter链式调用
        worksheet->rangeFormatter("A2:C2")
            .bold()
            .backgroundColor(Color::BLUE)
            .fontColor(Color::WHITE)
            .centerAlign()
            .allBorders(BorderStyle::Medium)
            .apply();
        std::cout << "✓ 表头格式化完成\n";
        
        // 3.3 格式化数据区域 - 演示批量边框
        worksheet->rangeFormatter("A3:C5")
            .allBorders(BorderStyle::Thin)
            .vcenterAlign()
            .apply();
        std::cout << "✓ 数据区域格式化完成\n";
        
        // 3.4 突出显示重要信息
        QuickFormat::highlight(*worksheet, "A5:C5", Color::YELLOW);
        std::cout << "✓ 重要信息突出显示完成\n";
        
        // 4. 添加使用提示
        worksheet->setValue(7, 0, "💡 使用提示:");
        worksheet->setValue(8, 0, "1. 使用rangeFormatter()进行批量格式化");
        worksheet->setValue(9, 0, "2. 使用QuickFormat快速应用常用样式");
        worksheet->setValue(10, 0, "3. 支持链式调用，代码更简洁");
        worksheet->setValue(11, 0, "4. 内部自动优化，性能更好");
        
        // 格式化提示文本
        QuickFormat::formatAsComment(*worksheet, "A7:A11");
        
        // 5. 保存文件
        workbook->save();
        std::cout << "\n✅ API使用指南创建完成！\n";
        std::cout << "文件已保存为: api_guide_demo.xlsx\n\n";
        
        // 6. 显示代码示例
        std::cout << "🔥 代码示例:\n";
        std::cout << "```cpp\n";
        std::cout << "// 1. 批量格式化（链式调用）\n";
        std::cout << "worksheet->rangeFormatter(\"A1:C10\")\n";
        std::cout << "    .bold()\n";
        std::cout << "    .backgroundColor(Color::BLUE)\n";
        std::cout << "    .centerAlign()\n";
        std::cout << "    .allBorders(BorderStyle::Medium)\n";
        std::cout << "    .apply();\n\n";
        
        std::cout << "// 2. 快速格式化\n";
        std::cout << "QuickFormat::formatAsCurrency(worksheet, \"B2:B10\", \"¥\");\n";
        std::cout << "QuickFormat::formatAsTable(worksheet, \"A1:D10\");\n\n";
        
        std::cout << "// 3. 突出显示\n";
        std::cout << "QuickFormat::highlight(worksheet, \"A5:C5\", Color::YELLOW);\n";
        std::cout << "QuickFormat::formatAsSuccess(worksheet, \"D1:D1\");\n";
        std::cout << "```\n\n";
        
        std::cout << "🎯 主要特性:\n";
        std::cout << "• 🚀 性能优化: 自动FormatRepository管理\n";
        std::cout << "• 🔗 链式调用: 代码更简洁易读\n";
        std::cout << "• 📦 丰富API: 覆盖常用格式化需求\n";
        std::cout << "• 🛡️ 类型安全: 编译时错误检查\n";
        std::cout << "• 📚 向后兼容: 不影响现有代码\n\n";
        
        return 0;
        
    } catch (const std::exception& e) {
        std::cerr << "❌ 错误: " << e.what() << std::endl;
        return 1;
    }
}