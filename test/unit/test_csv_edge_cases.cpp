#include "fastexcel/core/CSVProcessor.hpp"
#include "fastexcel/core/Path.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/Worksheet.hpp"
#include <fstream>
#include <iostream>

using namespace fastexcel::core;

int main() {
    try {
        std::cout << "=== CSV边界情况测试 ===" << std::endl;
        
        // 1. 创建包含各种数据类型的CSV测试文件
        std::ofstream test_file("test_edge_cases.csv");
        test_file << "名称,年龄,薪资,是否全职,入职日期,备注\n";
        test_file << "张三, 25 ,3500.50,true,2023-01-15,\"正常员工\"\n";
        test_file << "李四,-1,0,false,2023/02/20,\"包含\"\"引号\"\"的备注\"\n";
        test_file << "王五,  30  ,-1500.75,YES,2023-03-10,包含,逗号的文本\n";
        test_file << "赵六,abc,不是数字,1,2023-04-01,\n";
        test_file << "\"包含逗号,的姓名\",40,5000,N,2023-05-15,\"多行\\n文本\"\n";
        test_file << ",35,,true,,空值测试\n";
        test_file.close();
        
        std::cout << "✅ 创建测试CSV文件: test_edge_cases.csv" << std::endl;
        
        // 2. 测试CSV读取和类型推断
        auto workbook = Workbook::create("test_result.xlsx");
        auto worksheet = workbook->loadCSV("test_edge_cases.csv", "边界测试");
        
        if (!worksheet) {
            std::cout << "❌ CSV加载失败!" << std::endl;
            return 1;
        }
        
        std::cout << "✅ CSV加载成功" << std::endl;
        
        // 3. 验证数据类型推断
        auto [min_row, max_row, min_col, max_col] = worksheet->getUsedRangeFull();
        std::cout << "数据范围: (" << min_row << "," << min_col << ") -> (" 
                  << max_row << "," << max_col << ")" << std::endl;
        
        // 4. 检查特定单元格的类型推断结果
        std::cout << "\n=== 类型推断验证 ===" << std::endl;
        
        // 检查数字类型
        std::cout << "张三年龄 (B2): " << worksheet->getCell(1, 1).getValue<std::string>() << std::endl;
        std::cout << "李四年龄 (B3): " << worksheet->getCell(2, 1).getValue<std::string>() << std::endl;
        std::cout << "张三薪资 (C2): " << worksheet->getCell(1, 2).getValue<std::string>() << std::endl;
        std::cout << "王五薪资 (C4): " << worksheet->getCell(3, 2).getValue<std::string>() << std::endl;
        
        // 检查布尔类型
        std::cout << "张三是否全职 (D2): " << worksheet->getCell(1, 3).getValue<std::string>() << std::endl;
        std::cout << "李四是否全职 (D3): " << worksheet->getCell(2, 3).getValue<std::string>() << std::endl;
        std::cout << "王五是否全职 (D4): " << worksheet->getCell(3, 3).getValue<std::string>() << std::endl;
        
        // 5. 测试导出功能
        CSVOptions export_options;
        export_options.has_header = true;
        
        if (worksheet->saveAsCSV("test_export_result.csv", export_options)) {
            std::cout << "\n✅ CSV导出成功: test_export_result.csv" << std::endl;
        } else {
            std::cout << "\n❌ CSV导出失败!" << std::endl;
        }
        
        // 6. 保存Excel文件
        if (workbook->save()) {
            std::cout << "✅ Excel文件保存成功: test_result.xlsx" << std::endl;
        } else {
            std::cout << "❌ Excel文件保存失败!" << std::endl;
        }
        
        // 7. 测试不同分隔符的CSV
        std::ofstream semicolon_file("test_semicolon.csv");
        semicolon_file << "姓名;年龄;城市\n";
        semicolon_file << "Alice;25;北京\n";
        semicolon_file << "Bob;30;上海\n";
        semicolon_file.close();
        
        CSVOptions semicolon_options;
        semicolon_options.delimiter = ';';
        auto semicolon_sheet = workbook->loadCSV("test_semicolon.csv", "分号分隔", semicolon_options);
        
        if (semicolon_sheet) {
            std::cout << "✅ 分号分隔CSV加载成功" << std::endl;
        }
        
        std::cout << "\n=== 边界情况测试完成 ===" << std::endl;
        return 0;
        
    } catch (const std::exception& e) {
        std::cerr << "错误: " << e.what() << std::endl;
        return 1;
    }
}