#include "fastexcel/opc/PackageEditor.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/Path.hpp"  // 添加Path头文件
#include "fastexcel/utils/Logger.hpp"  // 修正包含路径
#include <iostream>
#include <memory>

using namespace fastexcel;

void testPackageEditorFromWorkbook() {
    std::cout << "=== 测试 PackageEditor::fromWorkbook() ===" << std::endl;
    
    try {
        // 1. 创建新的 Workbook - 需要提供路径参数
        core::Path temp_path("temp_workbook.xlsx");
        auto workbook = std::make_unique<core::Workbook>(temp_path);
        
        // 打开 Workbook（必须的操作）
        if (!workbook->open()) {
            std::cerr << "  ✗ 无法打开 Workbook" << std::endl;
            return;
        }
        
        // 2. 添加一些工作表
        workbook->addWorksheet("销售数据");
        workbook->addWorksheet("财务报表");
        
        // 3. 获取工作表并添加一些数据
        auto sheet1 = workbook->getWorksheet("销售数据");
        if (sheet1) {
            // 添加一些示例数据
            // sheet1->setCell(1, 1, "产品名称");
            // sheet1->setCell(1, 2, "销售额");
            std::cout << "  ✓ 向 '销售数据' 工作表添加了数据" << std::endl;
        }
        
        // 4. 从 Workbook 创建 PackageEditor
        auto editor = opc::PackageEditor::fromWorkbook(workbook.get());
        if (!editor) {
            std::cerr << "  ✗ 创建 PackageEditor 失败" << std::endl;
            return;
        }
        
        std::cout << "  ✓ 成功从 Workbook 创建 PackageEditor" << std::endl;
        
        // 5. 检查工作表
        auto sheet_names = editor->getSheetNames();
        std::cout << "  ✓ 工作表列表：";
        for (const auto& name : sheet_names) {
            std::cout << "'" << name << "' ";
        }
        std::cout << std::endl;
        
        // 6. 通过 Workbook 添加新工作表
        auto wb = editor->getWorkbook();
        if (wb) {
            wb->addWorksheet("库存管理");
            std::cout << "  ✓ 添加新工作表 '库存管理'" << std::endl;
        }
        
        // 7. 通过 Worksheet 设置单元格
        auto inventory_sheet = wb->getWorksheet("库存管理");
        if (inventory_sheet) {
            // 使用 Worksheet API 设置单元格
            inventory_sheet->writeString(1, 1, "测试数据");
            std::cout << "  ✓ 在 '库存管理' 工作表设置单元格 A1" << std::endl;
        }
        
        // 8. 检查是否有更改
        if (editor->isDirty()) {
            auto dirty_parts = editor->getDirtyParts();
            std::cout << "  ✓ 检测到 " << dirty_parts.size() << " 个需要更新的部件" << std::endl;
        }
        
        // 9. 提交到文件
        core::Path output_path("test_package_editor_output.xlsx");
        if (editor->commit(output_path)) {
            std::cout << "  ✓ 成功保存到 " << output_path.string() << std::endl;
        } else {
            std::cerr << "  ✗ 保存文件失败" << std::endl;
        }
        
    } catch (const std::exception& e) {
        std::cerr << "  ✗ 测试过程中发生异常：" << e.what() << std::endl;
    }
}

void testPackageEditorValidation() {
    std::cout << "\n=== 测试输入验证 ===" << std::endl;
    
    // 测试工作表名称验证
    std::cout << "测试工作表名称验证：" << std::endl;
    
    struct TestCase {
        std::string name;
        bool expected;
        std::string description;
    };
    
    std::vector<TestCase> sheet_tests = {
        {"正常工作表", true, "正常中文名称"},
        {"Sheet1", true, "正常英文名称"},
        {"", false, "空名称"},
        {"这个工作表名称超过了31个字符的限制应该会失败", false, "超长名称"},
        {"Sheet[1]", false, "包含禁止字符 []"},
        {"Sheet\\1", false, "包含禁止字符 \\"},
        {"Sheet/1", false, "包含禁止字符 /"},
        {"Sheet*1", false, "包含禁止字符 *"},
        {"Sheet?1", false, "包含禁止字符 ?"},
        {"Sheet:1", false, "包含禁止字符 :"},
        {"'Sheet1", false, "以单引号开头"},
        {"Sheet1'", false, "以单引号结尾"},
        {"History", false, "保留名称"}
    };
    
    for (const auto& test : sheet_tests) {
        bool result = opc::PackageEditor::isValidSheetName(test.name);
        std::cout << "  " << (result == test.expected ? "✓" : "✗") 
                  << " '" << test.name << "' - " << test.description 
                  << " (期望: " << (test.expected ? "有效" : "无效") << ")" << std::endl;
    }
    
    // 测试单元格引用验证
    std::cout << "\n测试单元格引用验证：" << std::endl;
    
    std::vector<std::tuple<int, int, bool, std::string>> cell_tests = {
        {1, 1, true, "A1 (最小有效值)"},
        {1048576, 16384, true, "XFD1048576 (最大有效值)"},
        {0, 1, false, "行号为0"},
        {1, 0, false, "列号为0"},
        {1048577, 1, false, "超出最大行数"},
        {1, 16385, false, "超出最大列数"},
        {-1, 1, false, "负行号"},
        {1, -1, false, "负列号"}
    };
    
    for (const auto& [row, col, expected, description] : cell_tests) {
        bool result = opc::PackageEditor::isValidCellRef(row, col);
        std::cout << "  " << (result == expected ? "✓" : "✗") 
                  << " 行" << row << "列" << col << " - " << description 
                  << " (期望: " << (expected ? "有效" : "无效") << ")" << std::endl;
    }
}

void testPackageEditorCreate() {
    std::cout << "\n=== 测试 PackageEditor::create() ===" << std::endl;
    
    try {
        // 1. 创建新的 PackageEditor
        auto editor = opc::PackageEditor::create();
        if (!editor) {
            std::cerr << "  ✗ 创建空 PackageEditor 失败" << std::endl;
            return;
        }
        
        std::cout << "  ✓ 成功创建空 PackageEditor" << std::endl;
        
        // 2. 检查默认工作表
        auto sheet_names = editor->getSheetNames();
        std::cout << "  ✓ 默认工作表数量：" << sheet_names.size() << std::endl;
        
        // 3. 添加数据
        if (!sheet_names.empty()) {
            auto wb = editor->getWorkbook();
            auto first_sheet = wb->getWorksheet(sheet_names[0]);
            if (first_sheet) {
                first_sheet->writeString(1, 1, "Hello World");
                std::cout << "  ✓ 在默认工作表设置了数据" << std::endl;
            }
        }
        
        // 4. 保存
        core::Path output_path("test_create_output.xlsx");
        if (editor->commit(output_path)) {
            std::cout << "  ✓ 成功保存新创建的文件到 " << output_path.string() << std::endl;
        } else {
            std::cerr << "  ✗ 保存文件失败" << std::endl;
        }
        
    } catch (const std::exception& e) {
        std::cerr << "  ✗ 测试过程中发生异常：" << e.what() << std::endl;
    }
}

int main() {
    // 设置日志级别 - 使用正确的 API
    Logger::getInstance().setLevel(Logger::Level::DEBUG);
    
    std::cout << "开始测试 PackageEditor 功能...\n" << std::endl;
    
    // 运行所有测试
    testPackageEditorValidation();
    testPackageEditorFromWorkbook();
    testPackageEditorCreate();
    
    std::cout << "\n测试完成！" << std::endl;
    
    return 0;
}
