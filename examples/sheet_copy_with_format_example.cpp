/**
 * @file sheet_copy_with_format_example.cpp
 * @brief 复制指定工作表并保持格式的示例
 * 
 * 这个示例演示如何：
 * - 读取源Excel文件的第三个工作表（屏柜分项表）
 * - 复制所有单元格内容和格式
 * - 写入到新的Excel文件
 * - 测试格式写入功能是否正常
 */

#include "fastexcel/FastExcel.hpp"
#include <iostream>
#include <string>
#include <chrono>
#include <fstream>
#include <sstream>
#include <vector>
#include <set>
#include <algorithm>

using namespace fastexcel;
using namespace fastexcel::core;

/**
 * @brief 工作表复制器，包含格式复制功能
 */
class SheetCopyWithFormat {
private:
    Path source_file_;
    Path target_file_;
    
public:
    SheetCopyWithFormat(Path source_file, Path target_file) 
        : source_file_(std::move(source_file)), target_file_(std::move(target_file)) {}
    
    /**
     * @brief 执行复制操作
     * @return 是否成功
     */
    bool copySheet() {
        try {
            std::cout << "=== Sheet Copy with Format Test ===" << std::endl;
            std::cout << "Source: " << source_file_ << std::endl;
            std::cout << "Target: " << target_file_ << std::endl;
            
            // 检查源文件是否存在
            if (!source_file_.exists()) {
                std::cerr << "Error: Source file does not exist" << std::endl;
                return false;
            }
            
            // 加载源工作簿
            std::cout << "\\nStep 1: Loading source workbook..." << std::endl;
            auto source_workbook = Workbook::loadForEdit(source_file_);
            if (!source_workbook) {
                std::cerr << "Error: Failed to load source workbook" << std::endl;
                return false;
            }
            std::cout << "OK: Source workbook loaded with " << source_workbook->getWorksheetCount() << " worksheets" << std::endl;
            
            
            auto source_worksheet = source_workbook->getWorksheet(0);
            if (!source_worksheet) {
                std::cerr << "Error: Failed to get third worksheet" << std::endl;
                return false;
            }
            std::cout << "OK: Got worksheet '" << source_worksheet->getName() << "'" << std::endl;
            
            // 创建目标工作簿
            std::cout << "\\nStep 2: Creating target workbook..." << std::endl;
            auto target_workbook = Workbook::create(target_file_);
            if (!target_workbook) {
                std::cerr << "Error: Failed to create target workbook" << std::endl;
                return false;
            }
            
            // 打开工作簿以启用编辑操作
            if (!target_workbook->open()) {
                std::cerr << "Error: Failed to open target workbook" << std::endl;
                return false;
            }
            std::cout << "OK: Target workbook created" << std::endl;
            
            // 复制样式数据和主题
            std::cout << "\\nStep 3: Copying styles and theme..." << std::endl;
            target_workbook->copyStylesFrom(*source_workbook);
            std::cout << "OK: Styles and theme copied automatically" << std::endl;
            
            // 创建目标工作表（使用源工作表名称）
            auto target_worksheet = target_workbook->addWorksheet(source_worksheet->getName());
            if (!target_worksheet) {
                std::cerr << "Error: Failed to create target worksheet" << std::endl;
                return false;
            }
            std::cout << "OK: Target worksheet renamed to '" << target_worksheet->getName() << "'" << std::endl;
            
            // 获取源工作表的使用范围
            auto used_range = source_worksheet->getUsedRange();
            int max_row = used_range.first;
            int max_col = used_range.second;
            int min_row = 0;  // 从第一行开始
            int min_col = 0;  // 从第一列开始
            
            std::cout << "\\nStep 4: Copying cells from range (0,0) to (" 
                     << max_row << "," << max_col << ")..." << std::endl;
            
            int copied_cells = 0;
            int formatted_cells = 0;
            
            // 复制每个单元格的内容和格式
            for (int row = min_row; row <= max_row; ++row) {
                for (int col = min_col; col <= max_col; ++col) {
                    const auto& source_cell = source_worksheet->getCell(row, col);
                    auto& target_cell = target_worksheet->getCell(row, col);
                    
                    // 复制单元格值
                    switch (source_cell.getType()) {
                    case CellType::String: {
                        auto value = source_cell.getStringValue();
                        if (!value.empty()) {
                            target_cell.setValue(value);
                            copied_cells++;
                        }
                        break;
                    }
                    case CellType::Number: {
                        auto value = source_cell.getNumberValue();
                        target_cell.setValue(value);
                        copied_cells++;
                        break;
                    }
                    case CellType::Boolean: {
                        auto value = source_cell.getBooleanValue();
                        target_cell.setValue(value);
                        copied_cells++;
                        break;
                    }
                    case CellType::Date: {
                        auto value = source_cell.getNumberValue(); // 日期作为数字存储
                        target_cell.setValue(value);
                        copied_cells++;
                        break;
                    }
                    case CellType::Formula: {
                        auto formula = source_cell.getFormula();
                        if (!formula.empty()) {
                            target_cell.setFormula(formula);
                            copied_cells++;
                        }
                        break;
                    }
                    case CellType::Empty:
                    default:
                        // 空单元格或其他类型，不复制值但仍需复制格式
                        break;
                    }
                    
                    // 复制格式（对所有单元格都执行，包括空单元格）
                    auto source_format = source_cell.getFormatDescriptor();
                    if (source_format) {
                        target_cell.setFormat(source_format);
                        formatted_cells++;
                    }
                }
                
                // 每100行显示一次进度
                if ((row - min_row + 1) % 100 == 0) {
                    std::cout << "  Processed " << (row - min_row + 1) << " rows..." << std::endl;
                }
            }
            
            std::cout << "OK: Copied " << copied_cells << " cells with " << formatted_cells << " formatted cells" << std::endl;
            
            // 🔧 关键修复：复制列信息（宽度和格式）
            std::cout << "\nStep 4.5: Copying column information..." << std::endl;
            
            // 调试：显示源工作表的列信息总数
            const auto& source_column_info = source_worksheet->getColumnInfo();
            std::cout << "DEBUG: Source worksheet has " << source_column_info.size() << " column configurations" << std::endl;
            
            int copied_columns = 0;
            int copied_column_formats = 0;
            for (int col = min_col; col <= max_col; ++col) {
                // 复制列宽
                double col_width = source_worksheet->getColumnWidth(col);
                if (col_width != target_worksheet->getColumnWidth(col)) {
                    target_worksheet->setColumnWidth(col, col_width);
                    copied_columns++;
                }
                
                // 复制列格式ID
                int col_format_id = source_worksheet->getColumnFormatId(col);
                if (col_format_id >= 0) {
                    target_worksheet->setColumnFormatId(col, col_format_id);
                    copied_column_formats++;
                    std::cout << "DEBUG: Copied column " << col << " format ID: " << col_format_id << std::endl;
                }
                
                // 复制列隐藏状态
                if (source_worksheet->isColumnHidden(col)) {
                    target_worksheet->hideColumn(col);
                }
            }
            std::cout << "OK: Copied " << copied_columns << " column width configurations and " 
                     << copied_column_formats << " column format configurations" << std::endl;
            
            // 🔧 最终诊断：检查目标工作表保存前的列信息状态
            const auto& target_column_info = target_worksheet->getColumnInfo();
            std::cout << "🔧 FINAL DEBUG: Target worksheet column_info_ size before save: " << target_column_info.size() << std::endl;
            for (int i = 0; i < 9; ++i) {
                int format_id = target_worksheet->getColumnFormatId(i);
                if (format_id >= 0) {
                    std::cout << "🔧 Target column " << i << " has format ID: " << format_id << std::endl;
                }
            }
            
            // 保存目标工作簿
            std::cout << "\\nStep 5: Saving target workbook..." << std::endl;
            bool saved = target_workbook->save();
            if (!saved) {
                std::cerr << "Error: Failed to save target workbook" << std::endl;
                return false;
            }
            std::cout << "OK: Target workbook saved successfully" << std::endl;
            
            // 显示统计信息
            std::cout << "\\n=== Copy Statistics ===" << std::endl;
            std::cout << "Source range: " << (max_row - min_row + 1) << " rows x " 
                     << (max_col - min_col + 1) << " cols" << std::endl;
            std::cout << "Copied cells: " << copied_cells << std::endl;
            std::cout << "Formatted cells: " << formatted_cells << std::endl;
            
            auto target_stats = target_workbook->getStyleStats();
            std::cout << "Target format count: " << target_stats.unique_formats << std::endl;
            std::cout << "Deduplication ratio: " << std::fixed << std::setprecision(2) 
                     << target_stats.deduplication_ratio * 100 << "%" << std::endl;
            
            return true;
            
        } catch (const std::exception& e) {
            std::cerr << "Error: " << e.what() << std::endl;
            return false;
        }
    }
};

int main() {
    try {
        std::cout << "FastExcel Sheet Copy with Format Example" << std::endl;
        std::cout << "Testing format writing functionality" << std::endl;
        std::cout << "Version: 2.0.0 - Modern C++ Architecture" << std::endl;
        
        // 记录开始时间
        auto start_time = std::chrono::high_resolution_clock::now();
        
        // 定义文件路径
        Path source_file("./辅材处理-张玥 机房建设项目（2025-JW13-W1007）测试.xlsx");
        Path target_file("./屏柜分项表_复制.xlsx");
        
        // 创建复制器并执行复制
        SheetCopyWithFormat copier(source_file, target_file);
        bool success = copier.copySheet();
        
        // 记录结束时间
        auto end_time = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time);
        
        std::cout << "\\n=== Result ===" << std::endl;
        if (success) {
            std::cout << "Success: Sheet copy with format completed in " 
                     << duration.count() << "ms" << std::endl;
        } else {
            std::cout << "Failed: Sheet copy failed after " 
                     << duration.count() << "ms" << std::endl;
            return 1;
        }
        
        return 0;
        
    } catch (const std::exception& e) {
        std::cerr << "Fatal error: " << e.what() << std::endl;
        return 1;
    }
}
