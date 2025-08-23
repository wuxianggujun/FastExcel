/**
 * @file 05_read_write_edit_example.cpp
 * @brief FastExcel读写编辑功能综合示例
 * 
 * 演示如何使用FastExcel进行Excel文件的读取、编辑和保存操作
 */

#include "fastexcel/FastExcel.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <iostream>
#include <vector>
#include <chrono>

using namespace fastexcel;

void demonstrateBasicReadWrite() {
    std::cout << "\n=== 基本读写功能演示 ===" << std::endl;
    
    try {
        // 1. 创建新工作簿并写入数据（使用新API）
        auto workbook = core::Workbook::create("sample_data.xlsx");
        if (!workbook) {
            FASTEXCEL_LOG_ERROR("无法创建工作簿");
            return;
        }
        
        auto worksheet = workbook->addSheet("员工数据");
        
        // 写入表头（使用新的模板化API）
        worksheet->setValue(0, 0, std::string("姓名"));
        worksheet->setValue(0, 1, std::string("年龄"));
        worksheet->setValue(0, 2, std::string("部门"));
        worksheet->setValue(0, 3, std::string("薪资"));
        worksheet->setValue(0, 4, std::string("入职日期"));
        
        // 写入数据（使用新的模板化API）
        std::vector<std::vector<std::string>> employee_data = {
            {"张三", "28", "技术部", "12000", "2023-01-15"},
            {"李四", "32", "销售部", "10000", "2022-06-20"},
            {"王五", "25", "人事部", "8000", "2023-03-10"},
            {"赵六", "35", "财务部", "15000", "2021-12-01"},
            {"钱七", "29", "技术部", "13000", "2022-09-15"}
        };
        
        for (size_t i = 0; i < employee_data.size(); ++i) {
            for (size_t j = 0; j < employee_data[i].size(); ++j) {
                if (j == 1 || j == 3) { // 年龄和薪资作为数字
                    worksheet->setValue(static_cast<int>(i + 1), static_cast<int>(j), 
                                     std::stod(employee_data[i][j]));
                } else {
                    worksheet->setValue(static_cast<int>(i + 1), static_cast<int>(j), 
                                     employee_data[i][j]);
                }
            }
        }
        
        // 添加公式
        worksheet->setValue(6, 0, std::string("平均薪资"));
        worksheet->getCell(6, 3).setFormula("AVERAGE(D2:D6)");
        
        // 设置文档属性（使用新API）
        workbook->setDocumentProperties(
            "员工信息管理系统",
            "员工数据演示",
            "FastExcel示例",
            "FastExcel公司",
            "演示基本读写功能"
        );
        
        // 保存文件
        if (workbook->save()) {
            std::cout << "✓ 成功创建并保存文件: sample_data.xlsx" << std::endl;
        }
        
        workbook->close();
        
    } catch (const std::exception& e) {
        std::cerr << "创建文件时发生错误: " << e.what() << std::endl;
    }
}

void demonstrateFileReading() {
    std::cout << "\n=== 文件读取功能演示 ===" << std::endl;
    
    try {
        // 读取刚才创建的文件（使用新API）
        auto workbook = core::Workbook::openReadOnly("sample_data.xlsx");
        if (!workbook) {
            std::cerr << "无法打开文件进行读取" << std::endl;
            return;
        }
        
        // 获取工作表名称
        auto worksheet_names = workbook->getSheetNames();
        std::cout << "✓ 发现 " << worksheet_names.size() << " 个工作表:" << std::endl;
        for (const auto& name : worksheet_names) {
            std::cout << "  - " << name << std::endl;
        }
        
        // 获取元数据（使用新API）
        const auto& doc_props = workbook->getDocumentProperties();
        std::cout << "✓ 文档信息:" << std::endl;
        std::cout << "  标题: " << doc_props.title << std::endl;
        std::cout << "  作者: " << doc_props.author << std::endl;
        std::cout << "  主题: " << doc_props.subject << std::endl;
        
        // 读取第一个工作表
        if (!worksheet_names.empty()) {
            auto worksheet = workbook->getSheet(worksheet_names[0]);
            if (worksheet) {
                std::cout << "✓ 成功读取工作表: " << worksheet_names[0] << std::endl;
                std::cout << "  单元格数量: " << worksheet->getCellCount() << std::endl;
                
                auto [max_row, max_col] = worksheet->getUsedRange();
                std::cout << "  使用范围: " << max_row + 1 << " 行 x " << max_col + 1 << " 列" << std::endl;
                
                // 显示前几行数据（使用新的模板化API）
                std::cout << "  数据预览:" << std::endl;
                for (int row = 0; row <= std::min(max_row, 3); ++row) {
                    std::cout << "    ";
                    for (int col = 0; col <= max_col; ++col) {
                        if (worksheet->hasCellAt(row, col)) {
                            const auto& cell = worksheet->getCell(row, col);
                            if (cell.isString()) {
                                std::cout << cell.getValue<std::string>() << "\t";
                            } else if (cell.isNumber()) {
                                std::cout << cell.getValue<double>() << "\t";
                            } else {
                                std::cout << "[其他]\t";
                            }
                        } else {
                            std::cout << "[空]\t";
                        }
                    }
                    std::cout << std::endl;
                }
                
                // 演示新的范围读取功能
                if (max_row >= 2 && max_col >= 2) {
                    std::cout << "  范围读取演示 (A1:C3):" << std::endl;
                    try {
                        auto range_data = worksheet->getRange<std::string>(0, 0, 2, 2);
                        for (const auto& row_data : range_data) {
                            std::cout << "    ";
                            for (const auto& cell_value : row_data) {
                                std::cout << cell_value << "\t";
                            }
                            std::cout << std::endl;
                        }
                    } catch (const std::exception& e) {
                        std::cout << "    范围读取失败: " << e.what() << std::endl;
                    }
                }
            }
        }
        
        workbook->close();
        
    } catch (const std::exception& e) {
        std::cerr << "读取文件时发生错误: " << e.what() << std::endl;
    }
}

void demonstrateEditingFeatures() {
    std::cout << "\n=== 编辑功能演示 ===" << std::endl;
    
    try {
        // 使用新的统一API进行编辑
        auto workbook = core::Workbook::openEditable("sample_data.xlsx");
        if (!workbook) {
            std::cerr << "无法打开文件进行编辑" << std::endl;
            return;
        }
        
        auto worksheet = workbook->getSheet("员工数据");
        if (!worksheet) {
            std::cerr << "找不到工作表: 员工数据" << std::endl;
            return;
        }
        
        // 1. 编辑单元格值（使用新的模板化API）
        std::cout << "✓ 编辑单元格数据..." << std::endl;
        worksheet->setValue(1, 3, 13000.0); // 修改张三的薪资
        worksheet->setValue(2, 2, std::string("市场部")); // 修改李四的部门
        
        // 2. 查找并替换
        std::cout << "✓ 执行查找替换..." << std::endl;
        int replacements = worksheet->findAndReplace("技术部", "研发部", false, false);
        std::cout << "  替换了 " << replacements << " 处 '技术部' -> '研发部'" << std::endl;
        
        // 3. 添加新数据（使用新的便捷方法）
        std::cout << "✓ 添加新员工数据..." << std::endl;
        std::vector<std::string> new_employee = {"孙八", "26", "研发部", "11000", "2023-08-01"};
        int new_row = worksheet->appendRow(new_employee);
        std::cout << "  新员工添加到第 " << new_row + 1 << " 行" << std::endl;
        
        // 4. 复制单元格
        std::cout << "✓ 复制单元格..." << std::endl;
        worksheet->copyCell(new_row, 0, new_row + 1, 0, true); // 复制新员工姓名到下一行
        worksheet->setValue(new_row + 1, 0, std::string("周九"));
        worksheet->copyRange(new_row, 1, new_row, 4, new_row + 1, 1, true); // 复制其他信息
        worksheet->setValue(new_row + 1, 1, 24.0);   // 修改年龄
        worksheet->setValue(new_row + 1, 3, 9500.0); // 修改薪资
        
        // 5. 排序数据
        std::cout << "✓ 按薪资排序..." << std::endl;
        worksheet->sortRange(1, 0, new_row + 1, 4, 3, false, false); // 按薪资列降序排序
        
        // 6. 添加新工作表
        std::cout << "✓ 添加新工作表..." << std::endl;
        auto summary_sheet = workbook->addSheet("薪资统计");
        summary_sheet->setValue(0, 0, std::string("部门"));
        summary_sheet->setValue(0, 1, std::string("平均薪资"));
        summary_sheet->setValue(1, 0, std::string("研发部"));
        summary_sheet->getCell(1, 1).setFormula("AVERAGEIF(员工数据.C:C,\"研发部\",员工数据.D:D)");
        summary_sheet->setValue(2, 0, std::string("市场部"));
        summary_sheet->getCell(2, 1).setFormula("AVERAGEIF(员工数据.C:C,\"市场部\",员工数据.D:D)");
        
        // 7. 全局查找
        std::cout << "✓ 执行全局查找..." << std::endl;
        core::Workbook::FindReplaceOptions options;
        options.match_case = false;
        auto search_results = workbook->findAll("研发部", options);
        std::cout << "  找到 " << search_results.size() << " 个 '研发部' 的匹配项:" << std::endl;
        for (const auto& [sheet_name, row, col] : search_results) {
            std::cout << "    工作表: " << sheet_name << ", 位置: " 
                      << static_cast<char>('A' + col) << (row + 1) << std::endl;
        }
        
        // 8. 获取统计信息
        auto stats = workbook->getStatistics();
        std::cout << "✓ 工作簿统计信息:" << std::endl;
        std::cout << "  工作表数量: " << stats.total_worksheets << std::endl;
        std::cout << "  总单元格数: " << stats.total_cells << std::endl;
        std::cout << "  格式数量: " << stats.total_formats << std::endl;
        std::cout << "  内存使用: " << stats.memory_usage / 1024 << " KB" << std::endl;
        
        // 保存编辑后的文件
        if (workbook->saveAs("edited_sample_data.xlsx")) {
            std::cout << "✓ 成功保存编辑后的文件: edited_sample_data.xlsx" << std::endl;
        }
        
    } catch (const std::exception& e) {
        std::cerr << "编辑文件时发生错误: " << e.what() << std::endl;
    }
}

void demonstrateAdvancedFeatures() {
    std::cout << "\n=== 高级功能演示 ===" << std::endl;
    
    try {
        // 创建一个复杂的工作簿
        auto workbook = core::Workbook::create("advanced_example.xlsx");
        if (!workbook) {
            std::cerr << "无法创建高级示例工作簿" << std::endl;
            return;
        }
        
        // 启用高性能模式
        workbook->setHighPerformanceMode(true);
        std::cout << "✓ 启用高性能模式" << std::endl;
        
        // 创建多个工作表
        auto sales_sheet = workbook->addSheet("销售数据");
        auto product_sheet = workbook->addSheet("产品信息");
        auto analysis_sheet = workbook->addSheet("数据分析");
        
        // 在销售数据表中添加大量数据
        std::cout << "✓ 生成大量测试数据..." << std::endl;
        sales_sheet->setValue(0, 0, std::string("日期"));
        sales_sheet->setValue(0, 1, std::string("产品"));
        sales_sheet->setValue(0, 2, std::string("销量"));
        sales_sheet->setValue(0, 3, std::string("单价"));
        sales_sheet->setValue(0, 4, std::string("总额"));
        
        // 生成1000行测试数据
        auto start_time = std::chrono::high_resolution_clock::now();
        
        for (int i = 1; i <= 1000; ++i) {
            sales_sheet->setValue(i, 0, std::string("2023-" + std::to_string((i % 12) + 1) + "-" + std::to_string((i % 28) + 1)));
            sales_sheet->setValue(i, 1, std::string("产品" + std::to_string((i % 10) + 1)));
            sales_sheet->setValue(i, 2, static_cast<double>((i % 100) + 1));
            sales_sheet->setValue(i, 3, 50.0 + (i % 200));
            sales_sheet->getCell(i, 4).setFormula("C" + std::to_string(i + 1) + "*D" + std::to_string(i + 1));
        }
        
        auto end_time = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time);
        std::cout << "  生成1000行数据耗时: " << duration.count() << "ms" << std::endl;
        
        // 设置自动筛选
        sales_sheet->setAutoFilter(0, 0, 1000, 4);
        std::cout << "✓ 设置自动筛选" << std::endl;
        
        // 冻结窗格
        sales_sheet->freezePanes(1, 0);
        std::cout << "✓ 冻结首行" << std::endl;
        
        // 在产品信息表中添加产品详情
        product_sheet->setValue(0, 0, std::string("产品编号"));
        product_sheet->setValue(0, 1, std::string("产品名称"));
        product_sheet->setValue(0, 2, std::string("类别"));
        product_sheet->setValue(0, 3, std::string("成本"));
        
        for (int i = 1; i <= 10; ++i) {
            product_sheet->setValue(i, 0, std::string("P" + std::to_string(i)));
            product_sheet->setValue(i, 1, std::string("产品" + std::to_string(i)));
            product_sheet->setValue(i, 2, std::string("类别" + std::to_string((i % 3) + 1)));
            product_sheet->setValue(i, 3, 20.0 + (i * 5));
        }
        
        // 在分析表中添加汇总信息
        analysis_sheet->setValue(0, 0, std::string("数据分析报告"));
        analysis_sheet->setValue(2, 0, std::string("总销售额"));
        analysis_sheet->getCell(2, 1).setFormula("SUM(销售数据.E:E)");
        analysis_sheet->setValue(3, 0, std::string("平均单价"));
        analysis_sheet->getCell(3, 1).setFormula("AVERAGE(销售数据.D:D)");
        analysis_sheet->setValue(4, 0, std::string("总销量"));
        analysis_sheet->getCell(4, 1).setFormula("SUM(销售数据.C:C)");
        
        // 合并单元格
        analysis_sheet->mergeCells(0, 0, 0, 3);
        std::cout << "✓ 合并标题单元格" << std::endl;
        
        // 设置工作表保护
        sales_sheet->protect("123456");
        std::cout << "✓ 保护销售数据工作表" << std::endl;
        
        // 设置文档属性
        workbook->setDocumentProperties(
            "销售数据分析系统",
            "大数据处理演示",
            "FastExcel高级示例",
            "FastExcel公司",
            "演示高级功能和大数据处理"
        );
        workbook->setProperty("版本", "1.0");
        workbook->setProperty("创建日期", "2023-08-04");
        
        // 保存文件
        start_time = std::chrono::high_resolution_clock::now();
        bool success = workbook->save();
        end_time = std::chrono::high_resolution_clock::now();
        duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time);
        
        if (success) {
            std::cout << "✓ 成功保存高级示例文件: advanced_example.xlsx" << std::endl;
            std::cout << "  保存耗时: " << duration.count() << "ms" << std::endl;
        }
        
        // 获取最终统计信息
        auto final_stats = workbook->getStatistics();
        std::cout << "✓ 最终统计信息:" << std::endl;
        std::cout << "  工作表数量: " << final_stats.total_worksheets << std::endl;
        std::cout << "  总单元格数: " << final_stats.total_cells << std::endl;
        std::cout << "  内存使用: " << final_stats.memory_usage / 1024 << " KB" << std::endl;
        
        workbook->close();
        
    } catch (const std::exception& e) {
        std::cerr << "高级功能演示时发生错误: " << e.what() << std::endl;
    }
}

int main() {
    std::cout << "FastExcel 读写编辑功能综合演示" << std::endl;
    std::cout << "版本: " << fastexcel::getVersion() << std::endl;
    
    try {
        // 初始化FastExcel库
        if (!fastexcel::initialize("logs/read_write_edit_example.log", true)) {
            std::cerr << "无法初始化FastExcel库" << std::endl;
            return -1;
        }
        
        // 演示各种功能
        demonstrateBasicReadWrite();
        demonstrateFileReading();
        demonstrateEditingFeatures();
        demonstrateAdvancedFeatures();
        
        std::cout << "\n=== 演示完成 ===" << std::endl;
        std::cout << "生成的文件:" << std::endl;
        std::cout << "  - sample_data.xlsx (基本示例)" << std::endl;
        std::cout << "  - edited_sample_data.xlsx (编辑后的文件)" << std::endl;
        std::cout << "  - advanced_example.xlsx (高级功能示例)" << std::endl;
        
        // 清理资源
        fastexcel::cleanup();
        
    } catch (const std::exception& e) {
        std::cerr << "程序执行时发生错误: " << e.what() << std::endl;
        return -1;
    }
    
    return 0;
}
