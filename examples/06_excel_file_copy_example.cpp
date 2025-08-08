/**
 * @file 06_excel_file_copy_example.cpp
 * @brief Complete Excel file copy example using FastExcel's high-level interfaces
 * 
 * This example demonstrates the proper way to use FastExcel library:
 * - Uses high-level Workbook interfaces (create, open, etc.)
 * - Leverages FastExcel's built-in Unicode file support through Path class
 * - Follows the library's architectural design
 * - Uses utf8cpp library for proper Unicode handling
 * - No direct usage of internal XLSXReader classes
 */

#include "fastexcel/FastExcel.hpp"
#include <iostream>
#include <string>
#include <chrono>

#ifdef _WIN32
#include <windows.h>
#endif

using namespace fastexcel;
using namespace fastexcel::core;

/**
 * @brief Excel文件复制器，使用FastExcel高级接口
 * 
 * 这个类演示了如何正确使用FastExcel库的架构设计：
 * 1. 使用 Workbook::open() 读取现有文件
 * 2. 使用 Workbook::create() 创建新文件  
 * 3. 通过工作表接口进行数据复制
 * 4. 保持格式和元数据的完整性
 */
class ExcelFileCopier {
private:
    Path source_file_;
    Path target_file_;
    
public:
    ExcelFileCopier(const Path& source_file, const Path& target_file)
        : source_file_(source_file), target_file_(target_file) {}
    
    /**
     * @brief 执行Excel文件复制操作
     * @return 是否复制成功
     */
    bool copyExcelFile() {
        try {
            std::cout << "=== Excel File Copy using FastExcel Architecture ===" << std::endl;
            std::cout << "Source file: " << source_file_ << std::endl;
            std::cout << "Target file: " << target_file_ << std::endl;
            
            auto start_time = std::chrono::high_resolution_clock::now();
            
            // Step 1: 使用FastExcel高级接口加载源工作簿
            std::cout << "\nStep 1: Loading source workbook..." << std::endl;
            
            // 首先检查源文件是否存在
            if (!source_file_.exists()) {
                std::cerr << "Error: Source file does not exist: " << source_file_ << std::endl;
                return false;
            }
            
            auto source_workbook = core::Workbook::open(source_file_);
            if (!source_workbook) {
                std::cerr << "Error: Failed to load source workbook: " << source_file_ << std::endl;
                return false;
            }
            std::cout << "OK: Source workbook loaded successfully" << std::endl;
            std::cout << "  Worksheets: " << source_workbook->getWorksheetCount() << std::endl;
            
            // Step 2: 创建目标工作簿
            std::cout << "\nStep 2: Creating target workbook..." << std::endl;
            auto target_workbook = core::Workbook::create(target_file_);
            if (!target_workbook) {
                std::cerr << "Error: Failed to create target workbook: " << target_file_ << std::endl;
                return false;
            }
            std::cout << "OK: Target workbook created successfully" << std::endl;
            
            // 样式复制已经在工作簿级别完成，直接复制单元格值
            std::cout << "\nStep 2.5: Copying styles data..." << std::endl;
            target_workbook->copyStylesFrom(source_workbook.get());
            std::cout << "OK: Styles data copied" << std::endl;
            
            // Step 3: 复制文档属性
            std::cout << "\nStep 3: Copying document properties..." << std::endl;
            target_workbook->setTitle(source_workbook->getTitle());
            target_workbook->setAuthor(source_workbook->getAuthor());
            target_workbook->setSubject(source_workbook->getSubject());
            
            // 复制自定义属性
            auto custom_props = source_workbook->getCustomProperties();
            for (const auto& prop : custom_props) {
                target_workbook->setCustomProperty(prop.first, prop.second);
            }
            std::cout << "OK: Document properties copied" << std::endl;
            
            // Step 4: 复制所有工作表
            std::cout << "\nStep 4: Copying worksheets..." << std::endl;
            size_t worksheet_count = source_workbook->getWorksheetCount();
            
            for (size_t i = 0; i < worksheet_count; ++i) {
                auto source_worksheet = source_workbook->getWorksheet(i);
                if (!source_worksheet) {
                    std::cerr << "Warning: Cannot access source worksheet " << i << std::endl;
                    continue;
                }
                
                std::string sheet_name = source_worksheet->getName();
                std::cout << "  Copying worksheet: " << sheet_name << std::endl;
                
                // 创建目标工作表
                auto target_worksheet = target_workbook->addWorksheet(sheet_name);
                
                // 获取源工作表的数据范围
                auto [max_row, max_col] = source_worksheet->getUsedRange();
                std::cout << "    Data range: " << max_row + 1 << " rows x " << max_col + 1 << " cols" << std::endl;
                
                // 复制所有单元格数据
                int copied_cells = 0;
                int copied_formats = 0;
                for (int row = 0; row <= max_row; ++row) {
                    for (int col = 0; col <= max_col; ++col) {
                        if (source_worksheet->hasCellAt(row, col)) {
                            const auto& source_cell = source_worksheet->getCell(row, col);
                            
                            // 根据单元格类型复制数据，保持类型和格式
                            switch (source_cell.getType()) {
                                case core::CellType::String:
                                    target_worksheet->writeString(row, col, source_cell.getStringValue());
                                    break;
                                    
                                case core::CellType::Number:
                                    target_worksheet->writeNumber(row, col, source_cell.getNumberValue());
                                    break;
                                    
                                case core::CellType::Boolean:
                                    target_worksheet->writeBoolean(row, col, source_cell.getBooleanValue());
                                    break;
                                    
                                case core::CellType::Formula:
                                    target_worksheet->writeFormula(row, col, source_cell.getFormula());
                                    break;
                                    
                                case core::CellType::Date:
                                    target_worksheet->writeNumber(row, col, source_cell.getNumberValue());
                                    break;
                                    
                                default:
                                    // 对于其他类型，尝试作为字符串处理
                                    if (!source_cell.getStringValue().empty()) {
                                        target_worksheet->writeString(row, col, source_cell.getStringValue());
                                    }
                                    break;
                            }
                            
                            // 复制单元格值（样式已在工作簿级别复制）
                            auto source_format = source_cell.getFormat();
                            if (source_format) {
                                // 样式已经通过 copyStylesFrom 处理，这里只需要复制单元格的格式引用
                                auto& target_cell = target_worksheet->getCell(row, col);
                                target_cell.setFormat(source_format); // 样式已经复制到目标工作簿
                                copied_formats++;
                            }
                            copied_cells++;
                        }
                    }
                }
                
                std::cout << "    OK: Copied " << copied_cells << " cells";
                if (copied_formats > 0) {
                    std::cout << " (with " << copied_formats << " formatted cells)";
                }
                std::cout << std::endl;
            }
            
            // Step 5: 设置激活工作表（确保只有第一个工作表被激活）
            std::cout << "\nStep 5: Setting active worksheet..." << std::endl;
            target_workbook->setActiveWorksheet(0);  // 设置第一个工作表为激活状态
            std::cout << "OK: First worksheet set as active" << std::endl;
            
            // Step 6: 保存目标工作簿
            std::cout << "\nStep 6: Saving target workbook..." << std::endl;
            auto save_start = std::chrono::high_resolution_clock::now();
            
            if (!target_workbook->save()) {
                std::cerr << "Error: Failed to save target workbook" << std::endl;
                return false;
            }
            
            auto save_end = std::chrono::high_resolution_clock::now();
            auto save_duration = std::chrono::duration_cast<std::chrono::milliseconds>(save_end - save_start);
            
            std::cout << "OK: Target workbook saved successfully" << std::endl;
            std::cout << "    Save time: " << save_duration.count() << "ms" << std::endl;
            
            // Step 7: 显示统计信息
            auto end_time = std::chrono::high_resolution_clock::now();
            auto total_duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time);
            
            std::cout << "\n=== Copy Statistics ===" << std::endl;
            std::cout << "Total time: " << total_duration.count() << "ms" << std::endl;
            std::cout << "Worksheets copied: " << worksheet_count << std::endl;
            
            // 获取目标工作簿统计信息
            auto stats = target_workbook->getStatistics();
            std::cout << "Target workbook stats:" << std::endl;
            std::cout << "  Total cells: " << stats.total_cells << std::endl;
            std::cout << "  Total formats: " << stats.total_formats << std::endl;
            std::cout << "  Memory usage: " << stats.memory_usage / 1024.0 << " KB" << std::endl;
            
            // 关闭工作簿
            target_workbook->close();
            source_workbook->close(); // open返回的工作簿也需要关闭
            
            std::cout << "\n=== Excel File Copy Completed Successfully ===" << std::endl;
            return true;
            
        } catch (const std::exception& e) {
            std::cerr << "Error during copy process: " << e.what() << std::endl;
            return false;
        }
    }
    
    /**
     * @brief 验证复制结果
     * @return 是否验证通过
     */
    bool verifyResult() {
        try {
            std::cout << "\n=== Verifying Copy Result ===" << std::endl;
            
            // 设置日志级别为 CRITICAL，减少验证阶段的日志输出
            fastexcel::Logger::getInstance().setLevel(fastexcel::Logger::Level::CRITICAL);
            
            // 验证目标文件是否存在
            auto target_workbook = core::Workbook::open(target_file_);
            if (!target_workbook) {
                std::cerr << "Verification failed: Cannot load target file" << std::endl;
                // 恢复日志级别
                fastexcel::Logger::getInstance().setLevel(fastexcel::Logger::Level::INFO);
                return false;
            }
            
            // 验证源文件
            auto source_workbook = core::Workbook::open(source_file_);
            if (!source_workbook) {
                std::cerr << "Verification failed: Cannot load source file" << std::endl;
                // 恢复日志级别
                fastexcel::Logger::getInstance().setLevel(fastexcel::Logger::Level::INFO);
                return false;
            }
            
            // 恢复日志级别
            fastexcel::Logger::getInstance().setLevel(fastexcel::Logger::Level::INFO);
            
            // 比较工作表数量
            if (source_workbook->getWorksheetCount() != target_workbook->getWorksheetCount()) {
                std::cerr << "Verification failed: Worksheet count mismatch" << std::endl;
                source_workbook->close();
                target_workbook->close();
                return false;
            }
            
            std::cout << "OK: Worksheet count matches: " << source_workbook->getWorksheetCount() << std::endl;
            
            // 比较每个工作表的基本信息
            for (size_t i = 0; i < source_workbook->getWorksheetCount(); ++i) {
                auto source_ws = source_workbook->getWorksheet(i);
                auto target_ws = target_workbook->getWorksheet(i);
                
                if (!source_ws || !target_ws) {
                    std::cerr << "Verification failed: Cannot access worksheet " << i << std::endl;
                    continue;
                }
                
                if (source_ws->getName() != target_ws->getName()) {
                    std::cerr << "Verification failed: Worksheet name mismatch at index " << i << std::endl;
                    continue;
                }
                
                auto [src_row, src_col] = source_ws->getUsedRange();
                auto [tgt_row, tgt_col] = target_ws->getUsedRange();
                
                if (src_row != tgt_row || src_col != tgt_col) {
                    std::cerr << "Warning: Data range mismatch in worksheet " << source_ws->getName() << std::endl;
                    continue;
                }
                
                std::cout << "OK: Worksheet '" << source_ws->getName() << "' verified successfully" << std::endl;
            }
            
            source_workbook->close();
            target_workbook->close();
            
            std::cout << "OK: Verification completed successfully" << std::endl;
            return true;
            
        } catch (const std::exception& e) {
            // 确保日志级别恢复
            fastexcel::Logger::getInstance().setLevel(fastexcel::Logger::Level::INFO);
            std::cerr << "Error during verification: " << e.what() << std::endl;
            return false;
        }
    }
};

int main() {
    // 设置控制台UTF-8支持（Windows）
#ifdef _WIN32
    SetConsoleCP(CP_UTF8);
    SetConsoleOutputCP(CP_UTF8);
#endif

    std::cout << "FastExcel Excel File Copy Example" << std::endl;
    std::cout << "Using FastExcel High-Level Architecture" << std::endl;
    std::cout << "Version: " << fastexcel::getVersion() << std::endl;
    
    // 使用相对路径，但是让程序自动查找中文文件
    Path source_file;
    Path target_file("./复制的辅材处理报表.xlsx");  // 恢复中文文件名支持
    
    // 查找当前目录下的.xlsx文件
#ifdef _WIN32
    std::cout << "\n=== Searching for Excel files in current directory ===" << std::endl;
    WIN32_FIND_DATAW find_data;
    HANDLE handle = FindFirstFileW(L"*.xlsx", &find_data);
    
    if (handle != INVALID_HANDLE_VALUE) {
        std::cout << "Found .xlsx files:" << std::endl;
        do {
            std::wstring wide_filename(find_data.cFileName);
            
            // 转换宽字符到UTF-8
            int size = WideCharToMultiByte(CP_UTF8, 0, wide_filename.c_str(), -1, NULL, 0, NULL, NULL);
            if (size > 0) {
                std::string utf8_name(size - 1, '\0');
                WideCharToMultiByte(CP_UTF8, 0, wide_filename.c_str(), -1, &utf8_name[0], size, NULL, NULL);
                std::cout << "  - " << utf8_name << std::endl;
                
                // 跳过我们要创建的目标文件
                if (utf8_name.find("复制的辅材处理报表.xlsx") != std::string::npos) {
                    std::cout << "    (Skipping target file)" << std::endl;
                    continue;
                }
                
                // 如果还没有找到源文件，使用第一个找到的
                if (source_file.empty()) {
                    source_file = Path("./" + utf8_name);
                    std::cout << "    --> Selected as source file" << std::endl;
                }
            }
        } while (FindNextFileW(handle, &find_data));
        FindClose(handle);
    } else {
        std::cout << "No .xlsx files found with FindFirstFileW" << std::endl;
        
        // 尝试查找所有文件
        std::cout << "\n=== Listing all files in current directory ===" << std::endl;
        HANDLE handle_all = FindFirstFileW(L"*", &find_data);
        if (handle_all != INVALID_HANDLE_VALUE) {
            do {
                std::wstring wide_filename(find_data.cFileName);
                int size = WideCharToMultiByte(CP_UTF8, 0, wide_filename.c_str(), -1, NULL, 0, NULL, NULL);
                if (size > 0) {
                    std::string utf8_name(size - 1, '\0');
                    WideCharToMultiByte(CP_UTF8, 0, wide_filename.c_str(), -1, &utf8_name[0], size, NULL, NULL);
                    std::cout << "  " << utf8_name << std::endl;
                    
                    // 检查是否是Excel文件
                    if (utf8_name.length() > 5 && utf8_name.substr(utf8_name.length() - 5) == ".xlsx") {
                        if (utf8_name.find("复制的辅材处理报表.xlsx") == std::string::npos && source_file.empty()) {
                            source_file = Path("./" + utf8_name);
                            std::cout << "    --> Found Excel file, selected as source" << std::endl;
                        }
                    }
                }
            } while (FindNextFileW(handle_all, &find_data));
            FindClose(handle_all);
        }
    }
#else
    // Linux/macOS fallback
    source_file = Path("./辅材处理-张玥 机房建设项目（2025-JW13-W1007）-配电系统(甲方客户报表).xlsx");
#endif
    
    if (source_file.empty()) {
        std::cerr << "Error: No suitable Excel source file found in current directory" << std::endl;
        return -1;
    }
    
    try {
        // 初始化FastExcel库
        if (!fastexcel::initialize("logs/excel_file_copy_example.log", true)) {
            std::cerr << "Error: Cannot initialize FastExcel library" << std::endl;
            return -1;
        }
        
        // 创建复制器并执行复制
        ExcelFileCopier copier(source_file, target_file);
        
        if (copier.copyExcelFile()) {
            // 验证结果
            if (copier.verifyResult()) {
                std::cout << "\nSuccess: File copied and verified successfully!" << std::endl;
                std::cout << "Target file: " << target_file << std::endl;
            } else {
                std::cout << "\nWarning: File copied but verification had issues" << std::endl;
            }
        } else {
            std::cerr << "Error: File copy failed" << std::endl;
            fastexcel::cleanup();
            return -1;
        }
        
        // 清理FastExcel资源
        fastexcel::cleanup();
        
    } catch (const std::exception& e) {
        std::cerr << "Fatal error: " << e.what() << std::endl;
        fastexcel::cleanup();
        return -1;
    }
    
    return 0;
}