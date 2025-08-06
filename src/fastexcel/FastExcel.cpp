#include "fastexcel/FastExcel.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/core/MemoryPool.hpp"
#include "fastexcel/core/WorkbookNew.hpp"
#include "fastexcel/core/StyleBuilder.hpp"

namespace fastexcel {

bool initialize(const std::string& log_file_path, bool enable_console) {
    // 初始化日志系统
    try {
        fastexcel::Logger::getInstance().initialize(log_file_path,
                                                   fastexcel::Logger::Level::DEBUG,
                                                   enable_console);
        return true;
    } catch (...) {
        return false;
    }
}

void initialize() {
    // 初始化日志系统（使用默认设置）
    try {
        fastexcel::Logger::getInstance().initialize("fastexcel.log",
                                                   fastexcel::Logger::Level::INFO,
                                                   true);
    } catch (...) {
        // 忽略日志初始化错误
    }
    
    // 初始化全局内存池等其他资源
    // 这里可以添加其他全局初始化代码
}

void cleanup() {
    // 清理内存池
    try {
        fastexcel::core::MemoryManager::getInstance().cleanup();
    } catch (...) {
        // 忽略清理错误
    }
    
    // 清理日志系统
    fastexcel::Logger::getInstance().shutdown();
}

// ========== 新架构 2.0 工厂函数实现 ==========

std::unique_ptr<core::NewWorkbook> createWorkbook(const std::string& filename) {
    return std::make_unique<core::NewWorkbook>(filename);
}

std::unique_ptr<core::NewWorkbook> openWorkbook(const std::string& filename) {
    return core::NewWorkbook::open(filename);
}

core::StyleBuilder createStyle() {
    return core::StyleBuilder();
}

core::NamedStyle createNamedStyle(const std::string& name, const core::StyleBuilder& builder) {
    return core::NamedStyle(name, builder);
}

// ========== 预定义样式工厂实现 ==========

namespace styles {

core::StyleBuilder title() {
    return core::StyleBuilder()
        .font("Arial", 18.0, true)
        .horizontalAlign(core::HorizontalAlign::Center)
        .verticalAlign(core::VerticalAlign::Center);
}

core::StyleBuilder header() {
    return core::StyleBuilder()
        .font("Arial", 12.0, true)
        .fill(core::Color(0xD9EDF7U))  // 浅蓝色背景
        .border(core::BorderStyle::Thin)
        .horizontalAlign(core::HorizontalAlign::Center);
}

core::StyleBuilder money() {
    return core::StyleBuilder()
        .numberFormat("$#,##0.00")
        .horizontalAlign(core::HorizontalAlign::Right);
}

core::StyleBuilder percent() {
    return core::StyleBuilder()
        .numberFormat("0.00%")
        .horizontalAlign(core::HorizontalAlign::Right);
}

core::StyleBuilder date() {
    return core::StyleBuilder()
        .numberFormat("yyyy-mm-dd")
        .horizontalAlign(core::HorizontalAlign::Center);
}

core::StyleBuilder border(core::BorderStyle style, core::Color color) {
    return core::StyleBuilder()
        .border(style, color);
}

core::StyleBuilder fill(core::Color color) {
    return core::StyleBuilder()
        .fill(color);
}

core::StyleBuilder font(const std::string& name, double size, bool bold) {
    return core::StyleBuilder()
        .font(name, size, bold);
}

} // namespace styles

// ========== 功能特性检测实现 ==========

bool hasNewArchitectureSupport() {
    return true;  // 新架构始终可用
}

std::string getNewArchitectureFeatures() {
    return "FastExcel 2.0 New Architecture Features:\n"
           "- Immutable FormatDescriptor with value object pattern\n"
           "- Thread-safe FormatRepository with automatic deduplication\n"
           "- Fluent StyleBuilder API with method chaining\n"
           "- Cross-workbook StyleTransferContext for style copying\n"
           "- Separated XML serialization layer\n"
           "- Modern C++17 design with smart pointers and RAII\n"
           "- Comprehensive predefined style library\n"
           "- Full backward compatibility with legacy APIs";
}

// ========== 迁移指南静态成员实现 ==========

const char* MigrationGuide::OLD_API_EXAMPLE = 
R"(// 旧API示例
Workbook wb("example.xlsx");
Format* fmt = wb.getFormatPool().createFormat();
fmt->setFontBold(true);
fmt->setFontSize(12.0);
Worksheet* ws = wb.getWorksheet(0);
ws->writeCell(0, 0, "Hello", fmt);
wb.save();)";

const char* MigrationGuide::NEW_API_EXAMPLE = 
R"(// 新API示例  
auto wb = createWorkbook("example.xlsx");
auto style = createStyle()
    .font("Arial", 12.0, true)
    .fill(Color::LIGHT_BLUE);
    
auto ws = wb->getWorksheet(0);
ws->writeCell(0, 0, "Hello", style);
wb->save();)";

const char* MigrationGuide::MIGRATION_STEPS = 
R"(迁移步骤：
1. 使用 createWorkbook() 替代 Workbook 构造函数
2. 使用 StyleBuilder 替代 Format 指针
3. 使用方法链替代逐个设置属性
4. 使用智能指针替代原始指针
5. 使用新的样式传输API进行跨工作簿操作
6. 利用预定义样式提高开发效率)";

} // namespace fastexcel