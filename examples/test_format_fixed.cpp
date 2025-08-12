#include "fastexcel/FastExcel.hpp"
#include "fastexcel/core/Path.hpp"
#include <iostream>

int main() {
    try {
        // 创建工作簿
        fastexcel::core::Path outputPath("test_format_fixed.xlsx");
        auto workbook = fastexcel::core::Workbook::create(outputPath);
        auto worksheet = workbook->addSheet("格式测试");

        // ========== 方法1: 直接设置FormatDescriptor（简单但不优化） ==========
        std::cout << "测试方法1: 直接设置FormatDescriptor..." << std::endl;
        
        // 创建红色背景格式
        auto redFormat = workbook->createStyleBuilder()
            .backgroundColor(fastexcel::core::Color(255, 0, 0))  // 红色
            .border(fastexcel::core::BorderStyle::Thin)
            .bold(true)
            .build();

        // 应用到A1单元格
        worksheet->setValue(0, 0, "方法1-直接设置");
        worksheet->getCell(0, 0).setFormat(
            std::make_shared<const fastexcel::core::FormatDescriptor>(redFormat)
        );

        // ========== 方法2: 通过FormatRepository（推荐，有优化） ==========
        std::cout << "测试方法2: 通过FormatRepository..." << std::endl;
        
        // 创建蓝色背景格式并添加到仓库
        int blueStyleId = workbook->addStyle(
            workbook->createStyleBuilder()
                .backgroundColor(fastexcel::core::Color(0, 0, 255))  // 蓝色
                .border(fastexcel::core::BorderStyle::Thin)
                .bold(true)
        );
        
        std::cout << "蓝色样式ID: " << blueStyleId << std::endl;

        // 应用到B1单元格
        worksheet->setValue(0, 1, "方法2-仓库管理");
        auto blueFormatDescriptor = workbook->getStyle(blueStyleId);
        worksheet->getCell(0, 1).setFormat(blueFormatDescriptor);

        // ========== 测试相同格式的去重优化 ==========
        std::cout << "测试格式去重优化..." << std::endl;
        
        // 方法2: 创建相同的蓝色格式（应该返回相同ID）
        int blueStyleId2 = workbook->addStyle(
            workbook->createStyleBuilder()
                .backgroundColor(fastexcel::core::Color(0, 0, 255))  // 相同的蓝色
                .border(fastexcel::core::BorderStyle::Thin)
                .bold(true)
        );
        
        std::cout << "第二次添加相同蓝色样式ID: " << blueStyleId2 << std::endl;
        std::cout << "是否去重成功: " << (blueStyleId == blueStyleId2 ? "是" : "否") << std::endl;

        // 应用到C1单元格
        worksheet->setValue(0, 2, "方法2-去重测试");
        worksheet->getCell(0, 2).setFormat(workbook->getStyle(blueStyleId2));

        // ========== 创建绿色格式用于对比 ==========
        int greenStyleId = workbook->addStyle(
            workbook->createStyleBuilder()
                .backgroundColor(fastexcel::core::Color(0, 255, 0))  // 绿色
                .border(fastexcel::core::BorderStyle::Medium)
                .italic(true)
        );

        worksheet->setValue(1, 0, "绿色格式");
        worksheet->getCell(1, 0).setFormat(workbook->getStyle(greenStyleId));

        // 保存文件
        bool result = workbook->save();
        
        if (result) {
            std::cout << "✅ 测试文件创建成功: test_format_fixed.xlsx" << std::endl;
            std::cout << "请打开文件检查:" << std::endl;
            std::cout << "  A1: 红色背景 (方法1)" << std::endl;
            std::cout << "  B1: 蓝色背景 (方法2)" << std::endl;
            std::cout << "  C1: 蓝色背景 (方法2-去重)" << std::endl;
            std::cout << "  A2: 绿色背景 (方法2-不同格式)" << std::endl;
        } else {
            std::cout << "❌ 保存失败" << std::endl;
        }

        return result ? 0 : 1;

    } catch (const std::exception& e) {
        std::cout << "异常: " << e.what() << std::endl;
        return 1;
    }
}