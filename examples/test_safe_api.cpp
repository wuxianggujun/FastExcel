#include "../src/fastexcel/core/Workbook.hpp"
#include "../src/fastexcel/core/Worksheet.hpp"
#include <iostream>
#include <optional>

using namespace fastexcel;
using namespace fastexcel::core;

int main() {
    try {
        std::cout << "测试安全API接口..." << std::endl;
        
        // 创建工作簿
        Path outputPath("test_safe_api.xlsx");
        auto workbook = Workbook::create(outputPath);
        if (!workbook) {
            std::cerr << "无法创建工作簿" << std::endl;
            return 1;
        }
        
        auto worksheet = workbook->addSheet("测试工作表");
        if (!worksheet) {
            std::cerr << "无法创建工作表" << std::endl;
            return 1;
        }
        
        // 设置一些测试数据
        worksheet->setValue(0, 0, std::string("Hello"));
        worksheet->setValue(0, 1, 123.45);
        worksheet->setValue(1, 0, std::string("World"));
        
        // 测试安全的单元格值获取
        std::cout << "=== 测试安全的单元格值获取 ===" << std::endl;
        
        // 1. 测试存在的单元格
        auto value1 = worksheet->tryGetValue<std::string>(0, 0);
        if (value1.has_value()) {
            std::cout << "✓ A1 = " << value1.value() << std::endl;
        } else {
            std::cout << "✗ 无法获取A1的值" << std::endl;
        }
        
        auto value2 = worksheet->tryGetValue<double>(0, 1);
        if (value2.has_value()) {
            std::cout << "✓ B1 = " << value2.value() << std::endl;
        } else {
            std::cout << "✗ 无法获取B1的值" << std::endl;
        }
        
        // 2. 测试不存在的单元格
        auto value3 = worksheet->tryGetValue<std::string>(10, 10);
        if (!value3.has_value()) {
            std::cout << "✓ K11 不存在，安全返回nullopt" << std::endl;
        } else {
            std::cout << "✗ 意外获取到K11的值: " << value3.value() << std::endl;
        }
        
        // 3. 测试安全的使用范围获取
        std::cout << "=== 测试安全的使用范围获取 ===" << std::endl;
        auto range = worksheet->tryGetUsedRange();
        if (range.has_value()) {
            std::cout << "✓ 使用范围: 最大行=" << range.value().first 
                      << ", 最大列=" << range.value().second << std::endl;
        } else {
            std::cout << "✗ 无法获取使用范围" << std::endl;
        }
        
        // 4. 测试安全的列宽获取
        std::cout << "=== 测试安全的列宽/行高获取 ===" << std::endl;
        auto width = worksheet->tryGetColumnWidth(0);
        if (width.has_value()) {
            std::cout << "✓ 第一列宽度: " << width.value() << std::endl;
        } else {
            std::cout << "✗ 无法获取第一列宽度" << std::endl;
        }
        
        auto height = worksheet->tryGetRowHeight(0);
        if (height.has_value()) {
            std::cout << "✓ 第一行高度: " << height.value() << std::endl;
        } else {
            std::cout << "✗ 无法获取第一行高度" << std::endl;
        }
        
        // 5. 测试工作簿级别的安全API
        std::cout << "=== 测试工作簿级别的安全API ===" << std::endl;
        
        // 安全获取工作表
        auto safeSheet = workbook->tryGetSheet("测试工作表");
        if (safeSheet.has_value()) {
            std::cout << "✓ 安全获取到工作表: " << safeSheet.value()->getName() << std::endl;
        } else {
            std::cout << "✗ 无法安全获取工作表" << std::endl;
        }
        
        // 安全获取不存在的工作表
        auto invalidSheet = workbook->tryGetSheet("不存在的工作表");
        if (!invalidSheet.has_value()) {
            std::cout << "✓ 不存在的工作表安全返回nullopt" << std::endl;
        } else {
            std::cout << "✗ 意外获取到不存在的工作表" << std::endl;
        }
        
        // 安全获取跨工作表的值
        auto crossValue = workbook->tryGetValue<std::string>("测试工作表", 0, 0);
        if (crossValue.has_value()) {
            std::cout << "✓ 跨工作表安全获取值: " << crossValue.value() << std::endl;
        } else {
            std::cout << "✗ 无法跨工作表安全获取值" << std::endl;
        }
        
        // 安全设置跨工作表的值
        bool setResult = workbook->trySetValue("测试工作表", 2, 0, std::string("安全设置的值"));
        if (setResult) {
            std::cout << "✓ 跨工作表安全设置值成功" << std::endl;
            
            // 验证设置结果
            auto verifyValue = workbook->tryGetValue<std::string>("测试工作表", 2, 0);
            if (verifyValue.has_value()) {
                std::cout << "✓ 验证设置的值: " << verifyValue.value() << std::endl;
            }
        } else {
            std::cout << "✗ 跨工作表安全设置值失败" << std::endl;
        }
        
        // 保存文件
        workbook->save();
        std::cout << "✅ 安全API测试完成，文件已保存为: test_safe_api.xlsx" << std::endl;
        
        return 0;
    } catch (const std::exception& e) {
        std::cerr << "❌ 错误: " << e.what() << std::endl;
        return 1;
    }
}