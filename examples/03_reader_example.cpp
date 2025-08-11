#include "fastexcel/utils/ModuleLoggers.hpp"
/**
 * @file 03_reader_example.cpp
 * @brief FastExcel读取功能示例
 * 
 * 演示如何使用FastExcel读取Excel文件并提取数据
 */

#include "fastexcel/reader/XLSXReader.hpp"
#include "fastexcel/core/Workbook.hpp"
#include <iostream>
#include <iomanip>
#include <algorithm>

int main() {
    try {
        // 创建XLSX读取器
        fastexcel::reader::XLSXReader reader("test_input.xlsx");
        
        // 打开文件
        if (!reader.open()) {
            EXAMPLE_ERROR("无法打开Excel文件");
            return -1;
        }
        
        EXAMPLE_INFO("=== FastExcel读取功能演示 ===");
        
        // 获取工作表名称列表
        auto worksheet_names = reader.getSheetNames();
        EXAMPLE_INFO("发现 {} 个工作表:", worksheet_names.size());
        for (size_t i = 0; i < worksheet_names.size(); ++i) {
            EXAMPLE_INFO("  {}. {}", (i + 1), worksheet_names[i]);
        }
        
        // 获取文档元数据
        auto metadata = reader.getMetadata();
        EXAMPLE_INFO("=== 文档元数据 ===");
        if (!metadata.title.empty()) {
            EXAMPLE_INFO("标题: {}", metadata.title);
        }
        if (!metadata.author.empty()) {
            EXAMPLE_INFO("作者: {}", metadata.author);
        }
        if (!metadata.subject.empty()) {
            EXAMPLE_INFO("主题: {}", metadata.subject);
        }
        if (!metadata.company.empty()) {
            EXAMPLE_INFO("公司: {}", metadata.company);
        }
        
        // 获取定义名称
        auto defined_names = reader.getDefinedNames();
        if (!defined_names.empty()) {
            std::cout << "\n=== 定义名称 ===" << std::endl;
            for (const auto& name : defined_names) {
                std::cout << "  - " << name << std::endl;
            }
        }
        
        // 读取第一个工作表的数据
        if (!worksheet_names.empty()) {
            std::cout << "\n=== 读取工作表: " << worksheet_names[0] << " ===" << std::endl;
            
            auto worksheet = reader.loadWorksheet(worksheet_names[0]);
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
                            
                            switch (cell.getType()) {
                                case fastexcel::core::CellType::String:
                                    value = "\"" + cell.getStringValue() + "\"";
                                    break;
                                case fastexcel::core::CellType::Number:
                                    value = std::to_string(cell.getNumberValue());
                                    break;
                                case fastexcel::core::CellType::Boolean:
                                    value = cell.getBooleanValue() ? "TRUE" : "FALSE";
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
                
            } else {
                std::cerr << "无法加载工作表: " << worksheet_names[0] << std::endl;
            }
        }
        
        // 演示加载整个工作簿
        std::cout << "\n=== 加载整个工作簿 ===" << std::endl;
        auto workbook = reader.loadWorkbook();
        if (workbook) {
            std::cout << "工作簿加载成功!" << std::endl;
            std::cout << "包含 " << workbook->getSheetCount() << " 个工作表" << std::endl;
            
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
        } else {
            std::cerr << "无法加载工作簿" << std::endl;
        }
        
        // 关闭文件
        reader.close();
        std::cout << "\n文件已关闭" << std::endl;
        
    } catch (const std::exception& e) {
        std::cerr << "发生错误: " << e.what() << std::endl;
        return -1;
    }
    
    std::cout << "\n=== 读取演示完成 ===" << std::endl;
    return 0;
}