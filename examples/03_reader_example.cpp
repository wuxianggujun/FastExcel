#include "fastexcel/utils/ModuleLoggers.hpp"
/**
 * @file 03_reader_example.cpp
 * @brief FastExcel读取功能示例
 * 
 * 演示如何使用FastExcel读取Excel文件并提取数据
 */

#include "fastexcel/FastExcel.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/Path.hpp"
#include <iostream>
#include <iomanip>
#include <algorithm>

int main() {
    try {
        // 初始化FastExcel库
        if (!fastexcel::initialize("logs/reader_example.log", true)) {
            EXAMPLE_ERROR("无法初始化FastExcel库");
            return -1;
        }
        
        EXAMPLE_INFO("=== FastExcel读取功能演示 ===");
        
        // 使用新API打开Excel文件进行只读访问
        auto workbook = fastexcel::core::Workbook::openForReading(fastexcel::core::Path("test_input.xlsx"));
        if (!workbook) {
            EXAMPLE_ERROR("无法打开Excel文件");
            return -1;
        }
        
        // 获取工作表名称列表
        auto worksheet_names = workbook->getSheetNames();
        EXAMPLE_INFO("发现 {} 个工作表:", worksheet_names.size());
        for (size_t i = 0; i < worksheet_names.size(); ++i) {
            EXAMPLE_INFO("  {}. {}", (i + 1), worksheet_names[i]);
        }
        
        // 获取文档元数据（新API）
        const auto& doc_props = workbook->getDocumentProperties();
        EXAMPLE_INFO("=== 文档元数据 ===");
        if (!doc_props.title.empty()) {
            EXAMPLE_INFO("标题: {}", doc_props.title);
        }
        if (!doc_props.author.empty()) {
            EXAMPLE_INFO("作者: {}", doc_props.author);
        }
        if (!doc_props.subject.empty()) {
            EXAMPLE_INFO("主题: {}", doc_props.subject);
        }
        if (!doc_props.company.empty()) {
            EXAMPLE_INFO("公司: {}", doc_props.company);
        }
        
        // 读取第一个工作表的数据
        if (!worksheet_names.empty()) {
            std::cout << "\n=== 读取工作表: " << worksheet_names[0] << " ===" << std::endl;
            
            auto worksheet = workbook->getSheet(worksheet_names[0]);
            if (worksheet) {
                std::cout << "工作表加载成功!" << std::endl;
                
                // 获取使用范围
                auto [max_row, max_col] = worksheet->getUsedRange();
                std::cout << "数据范围: " << max_row + 1 << " 行 x " << max_col + 1 << " 列" << std::endl;
                
                // 显示前10行10列的数据
                std::cout << "\n前10行10列数据预览:" << std::endl;
                std::cout << std::setw(8) << "行\\列";
                for (int col = 0; col < std::min(10, max_col + 1); ++col) {
                    std::cout << std::setw(12) << ("Col" + std::to_string(col + 1));
                }
                std::cout << std::endl;
                
                for (int row = 0; row < std::min(10, max_row + 1); ++row) {
                    std::cout << std::setw(8) << ("Row" + std::to_string(row + 1));
                    
                    for (int col = 0; col < std::min(10, max_col + 1); ++col) {
                        if (worksheet->hasCellAt(row, col)) {
                            const auto& cell = worksheet->getCell(row, col);
                            std::string value;
                            
                            // 使用新的模板化API获取值
                            switch (cell.getType()) {
                                case fastexcel::core::CellType::String:
                                    value = "\"" + cell.getValue<std::string>() + "\"";
                                    break;
                                case fastexcel::core::CellType::Number:
                                    value = std::to_string(cell.getValue<double>());
                                    break;
                                case fastexcel::core::CellType::Boolean:
                                    value = cell.getValue<bool>() ? "TRUE" : "FALSE";
                                    break;
                                case fastexcel::core::CellType::Formula:
                                    value = "=" + cell.getFormula();
                                    break;
                                default:
                                    value = "(empty)";
                                    break;
                            }
                            
                            // 截断长字符串
                            if (value.length() > 10) {
                                value = value.substr(0, 7) + "...";
                            }
                            
                            std::cout << std::setw(12) << value;
                        } else {
                            std::cout << std::setw(12) << "(empty)";
                        }
                    }
                    std::cout << std::endl;
                }
                
                // 统计信息
                std::cout << "\n=== 统计信息 ===" << std::endl;
                std::cout << "总单元格数: " << worksheet->getCellCount() << std::endl;
                
                // 演示新的模板化范围读取API
                if (max_row >= 2 && max_col >= 2) {
                    std::cout << "\n=== 演示范围读取 ===" << std::endl;
                    try {
                        // 读取A1:C3范围的字符串数据
                        auto range_data = worksheet->getRange<std::string>(0, 0, 2, 2);
                        std::cout << "A1:C3范围数据:" << std::endl;
                        for (size_t r = 0; r < range_data.size(); ++r) {
                            for (size_t c = 0; c < range_data[r].size(); ++c) {
                                std::cout << std::setw(12) << range_data[r][c];
                            }
                            std::cout << std::endl;
                        }
                    } catch (const std::exception& e) {
                        std::cout << "范围读取失败: " << e.what() << std::endl;
                    }
                }
                
            } else {
                std::cerr << "无法加载工作表: " << worksheet_names[0] << std::endl;
            }
        }
        
        // 演示工作簿统计信息（新API）
        std::cout << "\n=== 工作簿统计信息 ===" << std::endl;
        auto stats = workbook->getStatistics();
        std::cout << "工作表数量: " << stats.total_worksheets << std::endl;
        std::cout << "总单元格数: " << stats.total_cells << std::endl;
        std::cout << "内存使用: " << stats.memory_usage / 1024 << " KB" << std::endl;
        
        // 显示每个工作表的基本信息
        for (size_t i = 0; i < workbook->getSheetCount(); ++i) {
            auto ws = workbook->getSheet(i);
            if (ws) {
                auto used_range = ws->getUsedRange();
                int rows = used_range.first;
                int cols = used_range.second;
                std::cout << "  " << ws->getName() << ": "
                          << rows + 1 << "行 x " << cols + 1 << "列, "
                          << ws->getCellCount() << "个单元格" << std::endl;
            }
        }
        
        // 演示新的便捷访问方法
        std::cout << "\n=== 演示便捷访问方法 ===" << std::endl;
        if (workbook->hasSheet("Sheet1")) {
            auto sheet = workbook->findSheet("Sheet1");
            if (sheet && sheet->hasCellAt(0, 0)) {
                // 使用新的安全访问方法
                auto safe_value = sheet->tryGetValue<std::string>(0, 0);
                if (safe_value.has_value()) {
                    std::cout << "A1单元格值: " << safe_value.value() << std::endl;
                }
                
                // 使用默认值方法
                auto value_or_default = sheet->getValueOr<std::string>(0, 0, "默认值");
                std::cout << "A1单元格值（带默认值）: " << value_or_default << std::endl;
            }
        }
        
        // 关闭工作簿
        workbook->close();
        std::cout << "\n工作簿已关闭" << std::endl;
        
        // 清理资源
        fastexcel::cleanup();
        
    } catch (const std::exception& e) {
        std::cerr << "发生错误: " << e.what() << std::endl;
        return -1;
    }
    
    std::cout << "\n=== 读取演示完成 ===" << std::endl;
    return 0;
}