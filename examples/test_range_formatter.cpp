#include "../src/fastexcel/core/Color.hpp"
#include "../src/fastexcel/core/RangeFormatter.hpp"
#include "../src/fastexcel/core/StyleBuilder.hpp"
#include "../src/fastexcel/core/Workbook.hpp"
#include "../src/fastexcel/core/Worksheet.hpp"
#include <iostream>

using namespace fastexcel;
using namespace fastexcel::core;

int main() {
    try {
        // 创建工作簿和工作表 - 使用正确的工厂方法
        core::Path outputPath("test_range_formatter.xlsx");
        auto workbook = Workbook::create(outputPath);
        if (!workbook) {
            std::cerr << "无法创建工作簿" << std::endl;
            return 1;
        }
        
        auto worksheet = workbook->addSheet("测试范围格式化");
        if (!worksheet) {
            std::cerr << "无法创建工作表" << std::endl;
            return 1;
        }
        
        // 添加一些测试数据
        worksheet->setValue(0, 0, "产品名称");
        worksheet->setValue(0, 1, "数量");
        worksheet->setValue(0, 2, "价格");
        worksheet->setValue(0, 3, "总计");
        
        for (int i = 1; i <= 5; ++i) {
            worksheet->setValue(i, 0, "产品" + std::to_string(i));
            worksheet->setValue(i, 1, i * 10);
            worksheet->setValue(i, 2, 99.99);
            worksheet->setValue(i, 3, i * 10 * 99.99);
        }
        
        // 测试范围格式化
        std::cout << "测试范围格式化API...\n";
        
        // 1. 格式化标题行 - 使用合适的颜色
        worksheet->rangeFormatter("A1:D1")
            .bold()
            .backgroundColor(Color::BLUE)  // 使用预定义的蓝色
            .centerAlign()
            .allBorders(BorderStyle::Medium)
            .apply();
        
        std::cout << "✓ 标题行格式化完成\n";
        
        // 2. 格式化数据区域
        int processed = worksheet->rangeFormatter(1, 0, 5, 3)
            .allBorders(BorderStyle::Thin)
            .apply();
        
        std::cout << "✓ 数据区域格式化完成，处理了 " << processed << " 个单元格\n";
        
        // 3. 格式化价格列 - 使用合适的颜色
        worksheet->rangeFormatter("C2:D6")
            .backgroundColor(Color::GREEN)  // 使用预定义的绿色
            .rightAlign()
            .apply();
        
        std::cout << "✓ 价格列格式化完成\n";
        
        // 4. 预览功能测试
        auto formatter = worksheet->rangeFormatter("A1:D6");
        std::cout << "预览信息:\n" << formatter.preview();
        
        // 保存文件
        workbook->save();
        std::cout << "✅ 测试完成，文件已保存为: test_range_formatter.xlsx\n";
        
        return 0;
    } catch (const std::exception& e) {
        std::cerr << "❌ 错误: " << e.what() << std::endl;
        return 1;
    }
}