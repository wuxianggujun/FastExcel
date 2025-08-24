#include "fastexcel/core/CSVProcessor.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/Path.hpp"
#include <iostream>
#include <string>

using namespace fastexcel::core;

int main() {
    try {
        std::cout << "=== FastExcel CSV功能测试 ===" << std::endl;
        
        // 1. 创建工作簿和工作表
        auto workbook = Workbook::create("test_workbook.xlsx");
        auto worksheet = workbook->addSheet("测试数据");
        
        // 2. 添加测试数据
        std::cout << "添加测试数据..." << std::endl;
        worksheet->setValue(0, 0, std::string("姓名"));
        worksheet->setValue(0, 1, std::string("年龄"));
        worksheet->setValue(0, 2, std::string("分数"));
        worksheet->setValue(0, 3, std::string("是否通过"));
        
        worksheet->setValue(1, 0, std::string("张三"));
        worksheet->setValue(1, 1, 25);
        worksheet->setValue(1, 2, 89.5);
        worksheet->setValue(1, 3, true);
        
        worksheet->setValue(2, 0, std::string("李四"));
        worksheet->setValue(2, 1, 30);
        worksheet->setValue(2, 2, 76.2);
        worksheet->setValue(2, 3, true);
        
        worksheet->setValue(3, 0, std::string("王五"));
        worksheet->setValue(3, 1, 22);
        worksheet->setValue(3, 2, 58.7);
        worksheet->setValue(3, 3, false);
        
        // 3. 测试工作表范围获取
        std::cout << "测试范围获取..." << std::endl;
        auto [min_row, max_row, min_col, max_col] = worksheet->getUsedRangeFull();
        std::cout << "使用范围: (" << min_row << "," << min_col << ") -> (" 
                  << max_row << "," << max_col << ")" << std::endl;
        
        // 🚀 重要：保存Excel工作簿
        std::cout << "保存Excel工作簿..." << std::endl;
        if (workbook->save()) {
            std::cout << "Excel工作簿保存成功: test_workbook.xlsx" << std::endl;
        } else {
            std::cout << "Excel工作簿保存失败!" << std::endl;
        }
        
        // 4. 测试CSV导出为字符串
        std::cout << "\n测试CSV导出..." << std::endl;
        CSVOptions options = CSVOptions::standard();
        options.has_header = true;
        options.delimiter = ',';
        
        std::string csv_content = worksheet->toCSVString(options);
        std::cout << "CSV内容:\n" << csv_content << std::endl;
        
        // 5. 测试CSV保存到文件
        std::cout << "测试CSV文件保存..." << std::endl;
        std::string csv_filepath = "test_output.csv";
        if (worksheet->saveAsCSV(csv_filepath, options)) {
            std::cout << "CSV文件保存成功: " << csv_filepath << std::endl;
        } else {
            std::cout << "CSV文件保存失败!" << std::endl;
        }
        
        // 6. 测试另一种CSV保存方式（通过工作簿导出）
        std::cout << "测试工作簿CSV导出..." << std::endl;
        if (workbook->exportSheetAsCSV(0, "test_output_workbook.csv", options)) {
            std::cout << "工作簿CSV导出成功: test_output_workbook.csv" << std::endl;
        } else {
            std::cout << "工作簿CSV导出失败!" << std::endl;
        }
        
        // 7. 测试CSV加载
        std::cout << "\n测试CSV加载..." << std::endl;
        auto new_workbook = Workbook::create("test_loaded.xlsx");
        auto loaded_sheet = new_workbook->loadCSV(csv_filepath, "加载的数据", options);
        
        if (loaded_sheet) {
            std::cout << "CSV加载成功，工作表名称: " << loaded_sheet->getName() << std::endl;
            
            // 保存工作簿，确保Excel文件格式正确
            if (new_workbook->save()) {
                std::cout << "工作簿保存成功: test_loaded.xlsx" << std::endl;
            } else {
                std::cout << "工作簿保存失败!" << std::endl;
            }
            
            // 验证加载的数据
            auto [loaded_min_row, loaded_max_row, loaded_min_col, loaded_max_col] = loaded_sheet->getUsedRangeFull();
            std::cout << "加载的数据范围: (" << loaded_min_row << "," << loaded_min_col 
                      << ") -> (" << loaded_max_row << "," << loaded_max_col << ")" << std::endl;
            
            // 显示第一行数据作为验证
            std::cout << "第一行数据: ";
            for (int col = loaded_min_col; col <= loaded_max_col; ++col) {
                std::cout << "\"" << loaded_sheet->getCellDisplayValue(loaded_min_row, col) << "\" ";
            }
            std::cout << std::endl;
        } else {
            std::cout << "CSV加载失败!" << std::endl;
        }
        
        std::cout << "\n=== CSV功能测试完成 ===" << std::endl;
        return 0;
        
    } catch (const std::exception& e) {
        std::cerr << "测试过程中发生错误: " << e.what() << std::endl;
        return 1;
    }
}