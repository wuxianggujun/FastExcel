/**
 * @file 08_sheet_copy_with_format_example.cpp
 * @brief 复制指定工作表并保持格式的示例（展示新API功能）
 * 
 * 这个示例演示如何：
 * - 使用新的模板化API进行单元格操作
 * - 使用Excel地址格式（如"A1", "B2"）访问单元格
 * - 使用链式调用API简化代码
 * - 使用范围操作API批量处理数据
 * - 使用跨工作表访问方法
 * - 使用安全访问方法（tryGetValue, getValueOr）
 * - 读取源Excel文件并保持格式复制
 * - 测试新的API与现有代码的兼容性
 */

#include "fastexcel/FastExcel.hpp"
#include "fastexcel/core/WorksheetChain.hpp"  // 🚀 新增：链式调用支持
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
            auto source_workbook = Workbook::openForEditing(source_file_);
            if (!source_workbook) {
                std::cerr << "Error: Failed to load source workbook" << std::endl;
                return false;
            }
            std::cout << "OK: Source workbook loaded with " << source_workbook->getSheetCount() << " worksheets" << std::endl;
            
            
            auto source_worksheet = source_workbook->getSheet(0);
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
            
            std::cout << "OK: Target workbook created and ready" << std::endl;
            
            // 复制样式数据和主题
            std::cout << "\\nStep 3: Copying styles and theme..." << std::endl;
            target_workbook->copyStylesFrom(*source_workbook);
            std::cout << "OK: Styles and theme copied automatically" << std::endl;
            
            // 创建目标工作表（使用源工作表名称）
            auto target_worksheet = target_workbook->addSheet(source_worksheet->getName());
            if (!target_worksheet) {
                std::cerr << "Error: Failed to create target worksheet" << std::endl;
                return false;
            }
            std::cout << "OK: Target worksheet renamed to '" << target_worksheet->getName() << "'" << std::endl;
            
            // 🚀 新功能演示：使用新的API设置一些示例数据
            std::cout << "\n=== 新API功能演示 ===" << std::endl;
            
            // 演示1：使用模板化的setValue方法
            target_worksheet->setValue("A1", std::string("FastExcel 新API演示"));
            target_worksheet->setValue("A2", std::string("模板化方法"));
            target_worksheet->setValue("B2", 123.45);
            target_worksheet->setValue("C2", true);
            std::cout << "✓ 使用模板化setValue方法设置了A1-C2的值" << std::endl;
            
            // 演示2：使用Excel地址格式
            target_worksheet->setValue("D1", std::string("Excel地址格式"));
            target_worksheet->setValue("D2", 2024);
            std::cout << "✓ 使用Excel地址格式设置了D1-D2的值" << std::endl;
            
            // 演示3：使用链式调用
            target_worksheet->chain()
                .setValue("A3", std::string("链式调用"))
                .setValue("B3", 999.99)
                .setValue("C3", false)
                .setColumnWidth(0, 20.0)
                .setRowHeight(2, 25.0);
            std::cout << "✓ 使用链式调用设置了A3-C3的值和格式" << std::endl;
            
            // 演示4：使用范围操作
            std::vector<std::vector<std::string>> range_data = {
                {"范围操作", "演示", "数据"},
                {"第二行", "测试", "内容"}
            };
            target_worksheet->setRange("A4:C5", range_data);
            std::cout << "✓ 使用范围操作设置了A4:C5的数据" << std::endl;
            
            // 演示5：使用Workbook的跨工作表访问
            target_workbook->setValue("Sheet1!F1", std::string("跨工作表访问"));
            target_workbook->setValue(0, 5, 1, 42.0);  // 通过索引访问
            std::cout << "✓ 演示了跨工作表的单元格访问方法" << std::endl;
            
            // 演示6：安全访问方法
            auto safe_value = target_worksheet->tryGetValue<std::string>("A1");
            if (safe_value.has_value()) {
                std::cout << "✓ 安全获取A1的值: " << safe_value.value() << std::endl;
            }
            
            auto default_value = target_worksheet->getValueOr<double>("Z99", 0.0);
            std::cout << "✓ 获取Z99的值或默认值: " << default_value << std::endl;
            
            std::cout << "=== 新API演示完成，开始复制源文件 ===" << std::endl;
            
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
                    
                    // 🚀 新API：使用模板化方法复制单元格值
                    switch (source_cell.getType()) {
                    case CellType::String: {
                        auto value = source_cell.getValue<std::string>();
                        if (!value.empty()) {
                            target_worksheet->setValue(row, col, value);  // 使用新的模板API
                            copied_cells++;
                        }
                        break;
                    }
                    case CellType::Number: {
                        auto value = source_cell.getValue<double>();
                        target_worksheet->setValue(row, col, value);  // 使用新的模板API
                        copied_cells++;
                        break;
                    }
                    case CellType::Boolean: {
                        auto value = source_cell.getValue<bool>();
                        target_worksheet->setValue(row, col, value);  // 使用新的模板API
                        copied_cells++;
                        break;
                    }
                    case CellType::Date: {
                        auto value = source_cell.getValue<double>(); // 日期作为数字存储
                        target_worksheet->setValue(row, col, value);  // 使用新的模板API
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
        // 初始化Logger并启用DEBUG级别
        fastexcel::Logger::getInstance().initialize("logs/fastexcel.log", fastexcel::Logger::Level::DEBUG, true);
        
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
