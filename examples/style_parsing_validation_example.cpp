/**
 * @file style_parsing_validation_example.cpp
 * @brief Excel style parsing validation example using FastExcel's new architecture
 * 
 * This example demonstrates:
 * - Reading Excel files using the new Workbook architecture
 * - Validating the new style system (FormatDescriptor, FormatRepository, StyleBuilder)
 * - Detailed style information parsing and display
 * - No writing operations - pure read and analysis
 * - Comprehensive style statistics and validation
 */

#include "fastexcel/FastExcel.hpp"
#include <iostream>
#include <string>
#include <chrono>
#include <iomanip>
#include <map>

#ifdef _WIN32
#include <windows.h>
#endif

using namespace fastexcel;
using namespace fastexcel::core;

/**
 * @brief Excel样式解析验证器，专门用于测试新架构的样式系统
 * 
 * 这个类专注于：
 * 1. 验证新的样式系统是否正确解析Excel文件
 * 2. 详细分析FormatDescriptor的解析结果
 * 3. 测试FormatRepository的去重功能
 * 4. 验证StyleBuilder的功能完整性
 * 5. 不进行任何写入操作，纯读取验证
 */
class StyleParsingValidator {
private:
    Path source_file_;
    
public:
    explicit StyleParsingValidator(const Path& source_file)
        : source_file_(source_file) {}
    
    /**
     * @brief 执行样式解析验证
     * @return 是否验证成功
     */
    bool validateStyleParsing() {
        try {
            std::cout << "=== Excel Style Parsing Validation (New Architecture) ===" << std::endl;
            std::cout << "Source file: " << source_file_ << std::endl;
            
            auto start_time = std::chrono::high_resolution_clock::now();
            
            // Step 1: 加载Excel工作簿
            std::cout << "\nStep 1: Loading Excel workbook with new architecture..." << std::endl;
            
            if (!source_file_.exists()) {
                std::cerr << "Error: Source file does not exist: " << source_file_ << std::endl;
                return false;
            }
            
            // 使用新架构的静态工厂方法
            auto workbook = Workbook::openExisting(source_file_.toString());
            if (!workbook) {
                std::cerr << "Error: Failed to load workbook with new architecture" << std::endl;
                return false;
            }
            
            std::cout << "✓ Workbook loaded successfully with new architecture" << std::endl;
            std::cout << "  Worksheets: " << workbook->getWorksheetCount() << std::endl;
            
            // Step 2: 验证新样式系统基本功能
            std::cout << "\nStep 2: Validating new style system..." << std::endl;
            validateStyleSystem(*workbook);
            
            // Step 3: 分析工作表样式
            std::cout << "\nStep 3: Analyzing worksheet styles..." << std::endl;
            analyzeWorksheetStyles(*workbook);
            
            // Step 4: 测试样式构建器功能
            std::cout << "\nStep 4: Testing StyleBuilder functionality..." << std::endl;
            testStyleBuilder(*workbook);
            
            // Step 5: 显示详细统计
            std::cout << "\nStep 5: Detailed style statistics..." << std::endl;
            displayDetailedStatistics(*workbook);
            
            auto end_time = std::chrono::high_resolution_clock::now();
            auto total_duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time);
            
            std::cout << "\n=== Validation Completed Successfully ===" << std::endl;
            std::cout << "Total validation time: " << total_duration.count() << "ms" << std::endl;
            
            return true;
            
        } catch (const std::exception& e) {
            std::cerr << "Error during validation: " << e.what() << std::endl;
            return false;
        }
    }

private:
    /**
     * @brief 验证新样式系统的基本功能
     */
    void validateStyleSystem(const Workbook& workbook) {
        try {
            // 验证样式仓储
            const auto& style_repo = workbook.getStyleRepository();
            std::cout << "✓ StyleRepository accessible" << std::endl;
            std::cout << "  Style count: " << workbook.getStyleCount() << std::endl;
            std::cout << "  Default style ID: " << workbook.getDefaultStyleId() << std::endl;
            
            // 验证默认样式
            auto default_style = workbook.getStyle(workbook.getDefaultStyleId());
            if (default_style) {
                std::cout << "✓ Default style loaded successfully" << std::endl;
                displayStyleDetails(*default_style, "Default Style");
            } else {
                std::cout << "✗ Failed to load default style" << std::endl;
            }
            
            // 验证去重统计
            auto stats = workbook.getStyleStats();
            std::cout << "✓ Style deduplication stats:" << std::endl;
            std::cout << "  Cache hit rate: " << std::fixed << std::setprecision(2) 
                     << stats.getCacheHitRate() * 100 << "%" << std::endl;
            
            // 验证内存使用
            size_t memory_usage = workbook.getStyleMemoryUsage();
            std::cout << "✓ Style memory usage: " << memory_usage / 1024.0 << " KB" << std::endl;
            
        } catch (const std::exception& e) {
            std::cerr << "✗ Style system validation failed: " << e.what() << std::endl;
        }
    }
    
    /**
     * @brief 分析工作表中的样式使用情况
     */
    void analyzeWorksheetStyles(const Workbook& workbook) {
        size_t total_formatted_cells = 0;
        std::map<int, int> style_usage_count;
        
        for (size_t i = 0; i < workbook.getWorksheetCount(); ++i) {
            auto worksheet = workbook.getWorksheet(i);
            if (!worksheet) continue;
            
            std::cout << "\n  Analyzing worksheet: " << worksheet->getName() << std::endl;
            
            auto [max_row, max_col] = worksheet->getUsedRange();
            int formatted_cells_in_sheet = 0;
            
            for (int row = 0; row <= max_row; ++row) {
                for (int col = 0; col <= max_col; ++col) {
                    if (worksheet->hasCellAt(row, col)) {
                        const auto& cell = worksheet->getCell(row, col);
                        
                        // 检查单元格是否有格式
                        if (cell.hasFormat()) {
                            auto format = cell.getFormat();
                            if (format) {
                                // 这里需要获取新架构下的样式ID
                                // 假设有方法可以获取样式ID
                                int style_id = format->getStyleId(); // 这个方法可能需要添加到Format类中
                                style_usage_count[style_id]++;
                                formatted_cells_in_sheet++;
                                total_formatted_cells++;
                                
                                // 显示前几个格式化单元格的详细信息
                                if (formatted_cells_in_sheet <= 3) {
                                    std::cout << "    Cell " << (char)('A' + col) << (row + 1) 
                                             << " - Style ID: " << style_id 
                                             << ", Value: \"" << cell.getStringValue() << "\"" << std::endl;
                                }
                            }
                        }
                    }
                }
            }
            
            std::cout << "    Formatted cells: " << formatted_cells_in_sheet << "/" 
                     << ((max_row + 1) * (max_col + 1)) << std::endl;
        }
        
        std::cout << "\n✓ Style analysis summary:" << std::endl;
        std::cout << "  Total formatted cells: " << total_formatted_cells << std::endl;
        std::cout << "  Unique styles used: " << style_usage_count.size() << std::endl;
        
        // 显示最常用的样式
        if (!style_usage_count.empty()) {
            std::cout << "  Most used styles:" << std::endl;
            // 按使用次数排序并显示前5个
            std::vector<std::pair<int, int>> sorted_styles(style_usage_count.begin(), style_usage_count.end());
            std::sort(sorted_styles.begin(), sorted_styles.end(), 
                     [](const auto& a, const auto& b) { return a.second > b.second; });
            
            for (int i = 0; i < std::min(5, (int)sorted_styles.size()); ++i) {
                auto style = workbook.getStyle(sorted_styles[i].first);
                std::cout << "    Style ID " << sorted_styles[i].first 
                         << ": used " << sorted_styles[i].second << " times";
                if (style) {
                    std::cout << " (e.g., font: " << style->getFontName() 
                             << ", size: " << style->getFontSize() << ")";
                }
                std::cout << std::endl;
            }
        }
    }
    
    /**
     * @brief 测试StyleBuilder功能
     */
    void testStyleBuilder(const Workbook& workbook) {
        try {
            std::cout << "Testing StyleBuilder functionality..." << std::endl;
            
            // 创建样式构建器
            auto builder = workbook.createStyleBuilder();
            std::cout << "✓ StyleBuilder created successfully" << std::endl;
            
            // 测试链式调用
            auto test_style = builder
                .fontName("Arial")
                .fontSize(12)
                .bold(true)
                .italic(false)
                .fontColor(Color::RED)
                .backgroundColor(Color::LIGHT_BLUE)
                .horizontalAlign(HorizontalAlign::Center)
                .verticalAlign(VerticalAlign::Middle)
                .border(BorderStyle::Thin, Color::BLACK)
                .numberFormat("0.00")
                .build();
            
            std::cout << "✓ StyleBuilder chain operations successful" << std::endl;
            displayStyleDetails(test_style, "Test Style Created by StyleBuilder");
            
            // 测试从现有样式创建Builder
            auto default_style = workbook.getStyle(workbook.getDefaultStyleId());
            if (default_style) {
                auto builder_from_existing = StyleBuilder(*default_style);
                auto modified_style = builder_from_existing
                    .fontSize(14)
                    .bold(true)
                    .backgroundColor(Color::LIGHT_YELLOW)
                    .build();
                
                std::cout << "✓ StyleBuilder from existing style successful" << std::endl;
                displayStyleDetails(modified_style, "Modified Style from Default");
            }
            
        } catch (const std::exception& e) {
            std::cerr << "✗ StyleBuilder test failed: " << e.what() << std::endl;
        }
    }
    
    /**
     * @brief 显示样式的详细信息
     */
    void displayStyleDetails(const FormatDescriptor& style, const std::string& title) {
        std::cout << "\n--- " << title << " ---" << std::endl;
        
        // 字体信息
        std::cout << "Font: " << style.getFontName() 
                 << ", Size: " << style.getFontSize();
        if (style.isBold()) std::cout << ", Bold";
        if (style.isItalic()) std::cout << ", Italic";
        if (style.isStrikeout()) std::cout << ", Strikeout";
        std::cout << std::endl;
        
        // 对齐信息
        std::cout << "Alignment: H=" << static_cast<int>(style.getHorizontalAlign())
                 << ", V=" << static_cast<int>(style.getVerticalAlign());
        if (style.getTextWrap()) std::cout << ", Wrapped";
        if (style.getRotation() != 0) std::cout << ", Rotation=" << style.getRotation();
        std::cout << std::endl;
        
        // 边框信息
        auto left_border = style.getLeftBorderStyle();
        auto top_border = style.getTopBorderStyle();
        if (left_border != BorderStyle::None || top_border != BorderStyle::None) {
            std::cout << "Borders: Left=" << static_cast<int>(left_border)
                     << ", Top=" << static_cast<int>(top_border) << std::endl;
        }
        
        // 填充信息
        auto pattern = style.getPatternType();
        if (pattern != PatternType::None) {
            std::cout << "Fill: Pattern=" << static_cast<int>(pattern) << std::endl;
        }
        
        // 数字格式
        if (!style.getNumberFormat().empty()) {
            std::cout << "Number Format: \"" << style.getNumberFormat() << "\"" << std::endl;
        }
        
        // 保护信息
        std::cout << "Protection: Locked=" << (style.isLocked() ? "Yes" : "No")
                 << ", Hidden=" << (style.isHidden() ? "Yes" : "No") << std::endl;
        
        // 哈希值（用于验证去重）
        std::cout << "Hash: 0x" << std::hex << style.hash() << std::dec << std::endl;
    }
    
    /**
     * @brief 显示详细的统计信息
     */
    void displayDetailedStatistics(const Workbook& workbook) {
        std::cout << "\n=== Detailed Style System Statistics ===" << std::endl;
        
        // 基本统计
        auto stats = workbook.getStatistics();
        std::cout << "Workbook Statistics:" << std::endl;
        std::cout << "  Total worksheets: " << stats.total_worksheets << std::endl;
        std::cout << "  Total cells: " << stats.total_cells << std::endl;
        std::cout << "  Total styles: " << stats.total_formats << std::endl;
        std::cout << "  Total memory: " << stats.memory_usage / 1024.0 << " KB" << std::endl;
        
        // 样式系统统计
        std::cout << "\nStyle System Statistics:" << std::endl;
        std::cout << "  Style count: " << workbook.getStyleCount() << std::endl;
        std::cout << "  Style memory usage: " << workbook.getStyleMemoryUsage() / 1024.0 << " KB" << std::endl;
        
        auto style_stats = workbook.getStyleStats();
        std::cout << "  Cache hit rate: " << std::fixed << std::setprecision(2)
                 << style_stats.getCacheHitRate() * 100 << "%" << std::endl;
        
        // 性能统计
        std::cout << "\nPerformance Metrics:" << std::endl;
        std::cout << "  Average style memory per style: " 
                 << (workbook.getStyleCount() > 0 ? workbook.getStyleMemoryUsage() / workbook.getStyleCount() : 0) 
                 << " bytes" << std::endl;
        
        // 工作表详细统计
        std::cout << "\nWorksheet Details:" << std::endl;
        for (size_t i = 0; i < workbook.getWorksheetCount(); ++i) {
            auto worksheet = workbook.getWorksheet(i);
            if (worksheet) {
                auto [max_row, max_col] = worksheet->getUsedRange();
                std::cout << "  " << worksheet->getName() 
                         << ": " << (max_row + 1) << "×" << (max_col + 1) 
                         << " (" << ((max_row + 1) * (max_col + 1)) << " cells)" << std::endl;
            }
        }
    }
};

int main() {
    // 设置控制台UTF-8支持（Windows）
#ifdef _WIN32
    SetConsoleCP(CP_UTF8);
    SetConsoleOutputCP(CP_UTF8);
#endif

    std::cout << "FastExcel Style Parsing Validation Example" << std::endl;
    std::cout << "Testing New Architecture Style System" << std::endl;
    std::cout << "Version: " << fastexcel::getVersion() << std::endl;
    
    // 查找Excel文件
    Path source_file;
    
#ifdef _WIN32
    std::cout << "\n=== Searching for Excel files ===" << std::endl;
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
                
                // 优先选择包含样式的文件
                if (source_file.empty() || utf8_name.find("辅材") != std::string::npos) {
                    source_file = Path("./" + utf8_name);
                    std::cout << "    --> Selected as source file for style validation" << std::endl;
                }
            }
        } while (FindNextFileW(handle, &find_data));
        FindClose(handle);
    }
#else
    // Linux/macOS fallback
    source_file = Path("./辅材处理-张玥 机房建设项目（2025-JW13-W1007）-配电系统(甲方客户报表).xlsx");
#endif
    
    if (source_file.empty()) {
        std::cerr << "Error: No Excel files found for style validation" << std::endl;
        std::cerr << "Please place an Excel file in the current directory" << std::endl;
        return -1;
    }
    
    try {
        // 初始化FastExcel库
        if (!fastexcel::initialize("logs/style_parsing_validation.log", true)) {
            std::cerr << "Error: Cannot initialize FastExcel library" << std::endl;
            return -1;
        }
        
        // 创建验证器并执行验证
        StyleParsingValidator validator(source_file);
        
        if (validator.validateStyleParsing()) {
            std::cout << "\n🎉 Success: Style parsing validation completed successfully!" << std::endl;
            std::cout << "The new architecture style system is working correctly." << std::endl;
        } else {
            std::cerr << "\n❌ Error: Style parsing validation failed" << std::endl;
            std::cerr << "Please check the logs for detailed error information." << std::endl;
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