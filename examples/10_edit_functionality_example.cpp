#include "fastexcel/utils/ModuleLoggers.hpp"
/**
 * @file 10_edit_functionality_example.cpp
 * @brief FastExcel编辑功能测试示例
 * 
 * 这个示例演示如何：
 * - 拷贝现有Excel文件
 * - 对拷贝的文件进行编辑
 * - 修改单元格内容和格式
 * - 保存修改后的文件
 * - 验证编辑功能的正确性
 */

#include "fastexcel/FastExcel.hpp"
#include <iostream>
#include <string>
#include <chrono>
#include <filesystem>

using namespace fastexcel;
using namespace fastexcel::core;

int main() {
    try {
        // 初始化FastExcel库
        fastexcel::initialize("logs/edit_test.log", true);
        
        // 设置源文件路径和目标文件路径
        std::string source_file = "辅材处理-张玥 机房建设项目（2025-JW13-W1007）测试.xlsx";
        std::string target_file = "编辑测试_副本.xlsx";
        
        EXAMPLE_INFO("=== FastExcel 编辑功能测试示例 ===");
        EXAMPLE_INFO("源文件: {}", source_file);
        EXAMPLE_INFO("目标文件: {}", target_file);
        
        // 步骤1: 检查源文件是否存在
        Path source_path(source_file);
        if (!source_path.exists()) {
            EXAMPLE_ERROR("错误：源文件不存在: {}", source_file);
            return -1;
        }
        
        // 步骤2: 拷贝源文件到目标位置
        EXAMPLE_INFO("1. 拷贝源文件...");
        Path target_path(target_file);
        try {
            if (target_path.exists()) {
                target_path.remove();
                EXAMPLE_INFO("   - 删除现有目标文件");
            }
            source_path.copyTo(target_path);
            EXAMPLE_INFO("   - 文件拷贝成功");
        } catch (const std::exception& e) {
            EXAMPLE_ERROR("文件拷贝失败: {}", e.what());
            return -1;
        }
        
        // 步骤3: 打开拷贝的文件进行编辑
        EXAMPLE_INFO("2. 打开文件进行编辑...");
        auto workbook = Workbook::open(target_path);
        if (!workbook) {
            EXAMPLE_ERROR("无法打开工作簿进行编辑");
            return -1;
        }
        
        // 强制设置为批量模式，确保使用正确的压缩
        workbook->getOptions().mode = WorkbookMode::BATCH;
        workbook->setCompressionLevel(6);  // 明确设置压缩级别
        EXAMPLE_INFO("   - 设置为批量模式，压缩级别: 6");
        
        EXAMPLE_INFO("   - 工作簿打开成功");
        EXAMPLE_INFO("   - 工作表数量: {}", workbook->getWorksheetCount());
        
        // 4. 修改一些单元格内容和格式，应用新的样式
        EXAMPLE_INFO("4. 开始修改单元格内容和格式");
        
        // 获取第一个工作表
        auto worksheet = workbook->getWorksheet(0);
        if (!worksheet) {
            EXAMPLE_ERROR("无法获取第一个工作表");
            return 1;
        }
        
        // 创建一个现代化的样式：蓝色背景，白色字体，居中对齐
        auto modern_style = StyleBuilder()
            .font("Arial", 12, true)                        // 字体：Arial，12号，粗体
            .fontColor(Color(static_cast<uint32_t>(0xFFFFFF)))  // 白色字体
            .fill(Color(static_cast<uint32_t>(0x4472C4)))       // 蓝色背景
            .centerAlign()                                       // 居中对齐
            .vcenterAlign()                                      // 垂直居中
            .border(BorderStyle::Thin)                           // 细边框
            .build();
        
        // 应用样式到A1单元格  
        auto cell_a1 = worksheet->getCell(0, 0);  // A1单元格
        cell_a1.setValue("现代化样式标题");
        cell_a1.setFormat(std::make_shared<const FormatDescriptor>(modern_style));
        EXAMPLE_INFO("应用现代样式到A1单元格");
        
        // 在B2单元格添加数字并应用货币格式
        auto cell_b2 = worksheet->getCell(1, 1);  // B2单元格
        cell_b2.setValue(12345.67);
        auto money_style = StyleBuilder()
            .currency()                              // 货币格式
            .rightAlign()                            // 右对齐
            .build();
        cell_b2.setFormat(std::make_shared<const FormatDescriptor>(money_style));
        EXAMPLE_INFO("B2单元格设置货币格式: 12345.67");
        
        // 在C3单元格添加百分比
        auto cell_c3 = worksheet->getCell(2, 2);  // C3单元格
        cell_c3.setValue(0.856);
        auto percent_style = StyleBuilder()
            .percentage()                            // 百分比格式
            .centerAlign()                           // 居中对齐
            .build();
        cell_c3.setFormat(std::make_shared<const FormatDescriptor>(percent_style));
        EXAMPLE_INFO("C3单元格设置百分比格式: 85.6%");
        
        // 步骤5: 保存修改
        EXAMPLE_INFO("5. 保存修改...");
        if (workbook->saveAs(target_file)) {
            EXAMPLE_INFO("   - 文件保存成功");
        } else {
            EXAMPLE_ERROR("   - 文件保存失败");
            return -1;
        }
        
        // 步骤6: 验证编辑结果
        EXAMPLE_INFO("6. 验证编辑结果...");
        
        // 重新打开文件验证
        auto verify_workbook = Workbook::open(target_path);
        if (verify_workbook) {
            auto verify_worksheet = verify_workbook->getWorksheet(0);
            if (verify_worksheet) {
                auto verify_a1 = verify_worksheet->getCell(0, 0);
                auto verify_b2 = verify_worksheet->getCell(1, 1);
                auto verify_c3 = verify_worksheet->getCell(2, 2);
                
                EXAMPLE_INFO("   - 验证A1值: \"{}\"", verify_a1.getStringValue());
                EXAMPLE_INFO("   - 验证B2值: {}", verify_b2.getNumberValue());
                EXAMPLE_INFO("   - 验证C3值: {}", verify_c3.getNumberValue());
                
                EXAMPLE_INFO("   - 验证成功：文件编辑功能正常工作");
            }
        } else {
            EXAMPLE_ERROR("   - 验证失败：无法重新打开文件");
        }
        
        EXAMPLE_INFO("=== 编辑功能测试完成 ===");
        EXAMPLE_INFO("编辑后的文件保存在: {}", target_file);
        
        return 0;
        
    } catch (const std::exception& e) {
        EXAMPLE_ERROR("异常: {}", e.what());
        return -1;
    }
}
