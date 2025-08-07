/**
 * @file excel_format_parser_example.cpp
 * @brief Excel格式解析示例，使用FastExcel的Workbook高级API
 * 
 * 这个示例演示如何使用FastExcel库的高级API来解析Excel文件格式：
 * - 使用Workbook::loadForEdit()加载Excel文件
 * - 通过Worksheet API读取单元格数据和格式
 * - 使用Cell API获取格式信息
 * - 通过Format API解析详细的样式属性
 */

#include "fastexcel/FastExcel.hpp"
#include <iostream>
#include <string>
#include <iomanip>
#include <utility>

using namespace fastexcel;
using namespace fastexcel::core;

/**
 * @brief Excel格式解析器，使用FastExcel的高级API
 */
class ExcelFormatParser {
private:
    Path file_path_;
    
public:
    ExcelFormatParser(Path  file_path) : file_path_(std::move(file_path)) {}
    
    /**
     * @brief 使用Workbook API解析Excel文件格式
     * @return 是否解析成功
     */
    bool parseExcelFormat() {
        try {
            std::cout << "=== Excel Format Parser using FastExcel API ===" << std::endl;
            std::cout << "Target file: " << file_path_ << std::endl;
            
            // 检查文件是否存在
            if (!file_path_.exists()) {
                std::cerr << "Error: File does not exist: " << file_path_ << std::endl;
                return false;
            }
            
            // 使用Workbook::loadForEdit()加载Excel文件
            std::cout << "\nLoading Excel file using Workbook API..." << std::endl;
            auto workbook = core::Workbook::loadForEdit(file_path_);
            if (!workbook) {
                std::cerr << "Error: Failed to load workbook" << std::endl;
                return false;
            }
            std::cout << "OK: Workbook loaded successfully" << std::endl;
            
            // 解析工作簿基本信息
            parseWorkbookInfo(workbook.get());
            
            // 解析所有工作表的格式
            parseAllWorksheets(workbook.get());
            
            // 关闭工作簿
            workbook->close();
            
            std::cout << "\n=== Excel Format Parsing Completed ===" << std::endl;
            return true;
            
        } catch (const std::exception& e) {
            std::cerr << "Error during Excel format parsing: " << e.what() << std::endl;
            return false;
        }
    }
    
private:
    /**
     * @brief 使用Workbook API解析工作簿信息
     */
    void parseWorkbookInfo(core::Workbook* workbook) {
        std::cout << "\n=== Workbook Information ===" << std::endl;
        
        std::cout << "Title: " << workbook->getTitle() << std::endl;
        std::cout << "Author: " << workbook->getAuthor() << std::endl;
        std::cout << "Subject: " << workbook->getSubject() << std::endl;
        std::cout << "Total worksheets: " << workbook->getWorksheetCount() << std::endl;
        
        // 获取统计信息
        auto stats = workbook->getStatistics();
        std::cout << "Statistics:" << std::endl;
        std::cout << "  Total cells: " << stats.total_cells << std::endl;
        std::cout << "  Total formats: " << stats.total_formats << std::endl;
        std::cout << "  Memory usage: " << std::fixed << std::setprecision(2) 
                  << stats.memory_usage / 1024.0 << " KB" << std::endl;
    }
    
    /**
     * @brief 使用Workbook和Worksheet API解析所有工作表
     */
    void parseAllWorksheets(core::Workbook* workbook) {
        std::cout << "\n=== Worksheets Format Analysis ===" << std::endl;
        
        size_t worksheet_count = workbook->getWorksheetCount();
        
        for (size_t i = 0; i < worksheet_count; ++i) {
            auto worksheet = workbook->getWorksheet(i);
            if (!worksheet) {
                std::cout << "Warning: Cannot access worksheet " << i << std::endl;
                continue;
            }
            
            std::cout << "\n--- Worksheet #" << i << " ---" << std::endl;
            parseWorksheetFormat(worksheet.get());
        }
    }
    
    /**
     * @brief 使用Worksheet API解析单个工作表格式
     */
    void parseWorksheetFormat(core::Worksheet* worksheet) {
        std::string sheet_name = worksheet->getName();
        std::cout << "Worksheet name: " << sheet_name << std::endl;
        
        // 使用Worksheet::getUsedRange()获取数据范围
        auto [max_row, max_col] = worksheet->getUsedRange();
        std::cout << "Used range: " << max_row + 1 << " rows x " << max_col + 1 << " cols" << std::endl;
        
        // 分析前5x5区域的单元格格式
        int analyze_rows = std::min(max_row + 1, 5);
        int analyze_cols = std::min(max_col + 1, 5);
        
        analyzeCellFormats(worksheet, analyze_rows, analyze_cols);
    }
    
    /**
     * @brief 使用Cell和Format API分析单元格格式
     */
    void analyzeCellFormats(core::Worksheet* worksheet, int max_rows, int max_cols) {
        std::cout << "Cell format analysis (first " << max_rows << "x" << max_cols << " area):" << std::endl;
        
        int formatted_cells = 0;
        int total_cells = 0;
        
        for (int row = 0; row < max_rows; ++row) {
            for (int col = 0; col < max_cols; ++col) {
                if (worksheet->hasCellAt(row, col)) {
                    total_cells++;
                    const auto& cell = worksheet->getCell(row, col);
                    
                    // 获取单元格基本信息
                    std::string type_name = getCellTypeName(cell.getType());
                    std::string cell_value = getCellValueAsString(cell);
                    
                    // 检查是否有格式
                    auto format = cell.getFormat();
                    if (format) {
                        formatted_cells++;
                        std::cout << "\nCell(" << row << "," << col << "):" << std::endl;
                        std::cout << "  Type: " << type_name << std::endl;
                        std::cout << "  Value: \"" << cell_value << "\"" << std::endl;
                        
                        // 使用Format API解析详细格式信息
                        analyzeFormatDetails(format.get());
                    } else {
                        std::cout << "Cell(" << row << "," << col << "): " 
                                  << type_name << " = \"" << cell_value << "\" [No format]" << std::endl;
                    }
                }
            }
        }
        
        std::cout << "\nSummary:" << std::endl;
        std::cout << "  Total cells analyzed: " << total_cells << std::endl;
        std::cout << "  Formatted cells: " << formatted_cells << std::endl;
        std::cout << "  Format coverage: " << std::fixed << std::setprecision(1)
                  << (total_cells > 0 ? (formatted_cells * 100.0 / total_cells) : 0) << "%" << std::endl;
    }
    
    /**
     * @brief 使用Format API分析格式详细信息
     */
    void analyzeFormatDetails(core::Format* format) {
        std::cout << "  Format details:" << std::endl;
        
        // 字体格式
        if (format->hasFont()) {
            std::cout << "    Font: " << format->getFontName() 
                      << ", Size: " << format->getFontSize();
            if (format->isBold()) std::cout << ", Bold";
            if (format->isItalic()) std::cout << ", Italic";
            if (format->isStrikeout()) std::cout << ", Strikeout";
            std::cout << std::endl;
            std::cout << "    Font Color: RGB(" 
                      << static_cast<int>((format->getFontColor() >> 16) & 0xFF) << ","
                      << static_cast<int>((format->getFontColor() >> 8) & 0xFF) << ","
                      << static_cast<int>(format->getFontColor() & 0xFF) << ")" << std::endl;
        }
        
        // 对齐格式
        if (format->hasAlignment()) {
            std::cout << "    Alignment: ";
            if (format->getHorizontalAlign() != HorizontalAlign::None) {
                std::cout << "H=" << static_cast<int>(format->getHorizontalAlign()) << " ";
            }
            if (format->getVerticalAlign() != VerticalAlign::Bottom) {
                std::cout << "V=" << static_cast<int>(format->getVerticalAlign()) << " ";
            }
            if (format->isTextWrap()) std::cout << "Wrap ";
            if (format->getRotation() != 0) std::cout << "Rotation=" << format->getRotation() << "° ";
            if (format->getIndent() > 0) std::cout << "Indent=" << static_cast<int>(format->getIndent()) << " ";
            std::cout << std::endl;
        }
        
        // 边框格式
        if (format->hasBorder()) {
            std::cout << "    Borders: ";
            if (format->getLeftBorder() != BorderStyle::None) std::cout << "L ";
            if (format->getRightBorder() != BorderStyle::None) std::cout << "R ";
            if (format->getTopBorder() != BorderStyle::None) std::cout << "T ";
            if (format->getBottomBorder() != BorderStyle::None) std::cout << "B ";
            if (format->getDiagonalBorder() != BorderStyle::None) std::cout << "Diag ";
            std::cout << std::endl;
        }
        
        // 填充格式
        if (format->hasFill()) {
            std::cout << "    Fill: Pattern=" << static_cast<int>(format->getPattern());
            if (format->getPattern() != PatternType::None) {
                std::cout << ", BG=RGB("
                          << static_cast<int>((format->getBackgroundColor() >> 16) & 0xFF) << ","
                          << static_cast<int>((format->getBackgroundColor() >> 8) & 0xFF) << ","
                          << static_cast<int>(format->getBackgroundColor() & 0xFF) << ")";
            }
            std::cout << std::endl;
        }
        
        // 数字格式
        if (!format->getNumberFormat().empty()) {
            std::cout << "    Number Format: \"" << format->getNumberFormat() 
                      << "\" (Index: " << format->getNumberFormatIndex() << ")" << std::endl;
        }
        
        // 保护设置
        if (format->hasProtection()) {
            std::cout << "    Protection: ";
            if (format->isLocked()) std::cout << "Locked ";
            if (format->isHidden()) std::cout << "Hidden ";
            std::cout << std::endl;
        }
    }
    
    /**
     * @brief 获取单元格类型名称
     */
    std::string getCellTypeName(core::CellType type) {
        switch (type) {
            case core::CellType::String: return "String";
            case core::CellType::Number: return "Number";
            case core::CellType::Boolean: return "Boolean";
            case core::CellType::Formula: return "Formula";
            case core::CellType::Date: return "Date";
            default: return "Unknown";
        }
    }
    
    /**
     * @brief 使用Cell API获取单元格值的字符串表示
     */
    std::string getCellValueAsString(const core::Cell& cell) {
        switch (cell.getType()) {
            case core::CellType::String:
                return cell.getStringValue();
            case core::CellType::Number:
                return std::to_string(cell.getNumberValue());
            case core::CellType::Boolean:
                return cell.getBooleanValue() ? "TRUE" : "FALSE";
            case core::CellType::Formula:
                return cell.getFormula();
            default:
                return cell.getStringValue();
        }
    }
};

int main() {
    std::cout << "FastExcel Format Parser Example" << std::endl;
    std::cout << "Using Workbook, Worksheet, Cell and Format APIs" << std::endl;
    std::cout << "Version: " << fastexcel::getVersion() << std::endl;
    
    try {
        // 初始化FastExcel库
        if (!fastexcel::initialize("logs/excel_format_parser_example.log", true)) {
            std::cerr << "Error: Cannot initialize FastExcel library" << std::endl;
            return -1;
        }
        
        // 创建格式解析器并执行解析
        ExcelFormatParser parser(Path("./辅材处理-张玥 机房建设项目（2025-JW13-W1007）-配电系统(甲方客户报表).xlsx"));
        
        if (parser.parseExcelFormat()) {
            std::cout << "\nSuccess: Excel format parsing completed!" << std::endl;
        } else {
            std::cerr << "Error: Excel format parsing failed" << std::endl;
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
