/**
 * @file simple_edit_example.cpp
 * @brief 简化的Excel文件编辑示例
 * 
 * 展示如何直接打开XLSX文件进行编辑，无需复杂的API调用
 */

#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/Exception.hpp"
#include <iostream>

using namespace fastexcel;

int main() {
    try {
        // 方法1: 直接打开现有文件进行编辑
        std::cout << "=== 直接编辑现有文件 ===" << std::endl;
        
        // 使用loadForEdit直接加载文件进行编辑
        auto workbook = core::Workbook::loadForEdit("data.xlsx");
        if (!workbook) {
            std::cout << "文件不存在，创建新文件..." << std::endl;
            workbook = core::Workbook::create("data.xlsx");
            workbook->open();
        }
        
        // 获取或创建工作表
        auto sheet = workbook->getWorksheet("Sheet1");
        if (!sheet) {
            sheet = workbook->addWorksheet("Sheet1");
        }
        
        // 直接编辑数据
        sheet->writeString(0, 0, "姓名");
        sheet->writeString(0, 1, "年龄");
        sheet->writeString(0, 2, "部门");
        
        sheet->writeString(1, 0, "张三");
        sheet->writeNumber(1, 1, 25);
        sheet->writeString(1, 2, "技术部");
        
        sheet->writeString(2, 0, "李四");
        sheet->writeNumber(2, 1, 30);
        sheet->writeString(2, 2, "销售部");
        
        // 编辑现有数据
        sheet->editCellValue(1, 1, 26.0); // 修改张三的年龄
        
        // 查找和替换
        int replacements = sheet->findAndReplace("技术部", "研发部", false, false);
        std::cout << "替换了 " << replacements << " 个单元格" << std::endl;
        
        // 保存文件
        workbook->save();
        std::cout << "文件已保存" << std::endl;
        
        // 方法2: 一步式文件处理
        std::cout << "\n=== 一步式文件处理 ===" << std::endl;
        
        // 直接处理文件，无需手动管理工作簿生命周期
        auto processFile = [](const std::string& filename) {
            auto wb = core::Workbook::loadForEdit(filename);
            if (!wb) {
                wb = core::Workbook::create(filename);
                wb->open();
            }
            
            auto ws = wb->getWorksheet("数据");
            if (!ws) {
                ws = wb->addWorksheet("数据");
            }
            
            // 添加一些数据
            ws->writeString(0, 0, "产品");
            ws->writeString(0, 1, "价格");
            ws->writeString(1, 0, "苹果");
            ws->writeNumber(1, 1, 5.5);
            ws->writeString(2, 0, "香蕉");
            ws->writeNumber(2, 1, 3.2);
            
            // 自动保存
            wb->save();
            return wb->getWorksheetCount();
        };
        
        int sheetCount = processFile("products.xlsx");
        std::cout << "处理完成，共 " << sheetCount << " 个工作表" << std::endl;
        
        // 方法3: 批量编辑
        std::cout << "\n=== 批量编辑示例 ===" << std::endl;
        
        auto batchEdit = core::Workbook::loadForEdit("data.xlsx");
        if (batchEdit) {
            // 批量重命名工作表
            std::unordered_map<std::string, std::string> renameMap = {
                {"Sheet1", "员工信息"}
            };
            batchEdit->batchRenameWorksheets(renameMap);
            
            // 全局查找替换
            core::Workbook::FindReplaceOptions options;
            options.case_sensitive = false;
            int totalReplacements = batchEdit->findAndReplaceAll("销售部", "市场部", options);
            std::cout << "全局替换了 " << totalReplacements << " 个单元格" << std::endl;
            
            batchEdit->save();
        }
        
        // 方法4: 读取和显示数据
        std::cout << "\n=== 读取数据 ===" << std::endl;
        
        auto readWorkbook = core::Workbook::loadForEdit("data.xlsx");
        if (readWorkbook) {
            auto readSheet = readWorkbook->getWorksheet("员工信息");
            if (readSheet) {
                std::cout << "员工信息:" << std::endl;
                for (int row = 0; row < 3; ++row) {
                    for (int col = 0; col < 3; ++col) {
                        std::cout << readSheet->getCellString(row, col) << "\t";
                    }
                    std::cout << std::endl;
                }
            }
        }
        
        std::cout << "\n所有操作完成！" << std::endl;
        
    } catch (const core::FastExcelException& e) {
        std::cerr << "错误: " << e.getDetailedMessage() << std::endl;
        return 1;
    } catch (const std::exception& e) {
        std::cerr << "系统错误: " << e.what() << std::endl;
        return 1;
    }
    
    return 0;
}