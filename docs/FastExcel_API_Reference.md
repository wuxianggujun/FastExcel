# FastExcel API 参考文档

本文档提供了 FastExcel 库的最新 API 参考，基于当前代码实现编写，包括所有类、方法、使用示例和详细说明。

## 目录

1. [核心架构概述](#核心架构概述)
2. [快速开始](#快速开始)
3. [Workbook 类](#workbook-类)
4. [Worksheet 类](#worksheet-类)
5. [样式系统](#样式系统)
6. [Cell 类](#cell-类)
7. [流式处理](#流式处理)
8. [文档属性管理](#文档属性管理)
9. [工具类](#工具类)
10. [最佳实践](#最佳实践)

## 核心架构概述

FastExcel 采用现代 C++17 设计，核心组件包括：

```
┌─────────────────────────────────────────┐
│            FastExcel 架构                │
├─────────────────────────────────────────┤
│  应用层: 用户 API                        │
│  ├─ Workbook (工作簿管理)                │
│  ├─ Worksheet (工作表操作)               │
│  └─ FormatDescriptor & StyleBuilder      │
├─────────────────────────────────────────┤
│  核心层: 数据管理                        │
│  ├─ FormatRepository (格式仓储)          │
│  ├─ CustomPropertyManager               │
│  └─ DefinedNameManager                  │
├─────────────────────────────────────────┤
│  流式处理层                              │
│  ├─ XMLStreamWriter/Reader              │
│  ├─ BatchFileWriter                     │
│  └─ StreamingFileWriter                 │
├─────────────────────────────────────────┤
│  存储层: 文件管理                        │
│  └─ FileManager & ZipArchive            │
└─────────────────────────────────────────┘
```

### 关键设计特性

- **智能模式选择**: 根据数据量自动选择批量或流式处理
- **格式仓储去重**: 自动去除重复的格式定义，优化文件大小
- **回调式XML生成**: 统一的XML生成接口，支持批量和流式两种模式
- **内存管理优化**: 使用智能指针和RAII模式管理资源

## 快速开始

### 基本使用示例

```cpp
#include "fastexcel/FastExcel.hpp"
using namespace fastexcel::core;

int main() {
    // 创建工作簿
    auto workbook = Workbook::create("example.xlsx");
    workbook->open();
    
    // 添加工作表
    auto worksheet = workbook->addWorksheet("数据表");
    
    // 创建样式
    StyleBuilder header_style;
    header_style.font().bold(true).color(0xFFFFFF);
    header_style.fill().backgroundColor(0x4472C4);
    int header_style_id = workbook->addStyle(header_style);
    
    // 写入数据
    worksheet->writeString(0, 0, "姓名", header_style_id);
    worksheet->writeString(0, 1, "年龄", header_style_id);
    worksheet->writeString(1, 0, "张三");
    worksheet->writeNumber(1, 1, 25);
    
    // 保存文件
    workbook->save();
    
    return 0;
}
```

### 高性能模式

```cpp
// 启用高性能模式
workbook->setHighPerformanceMode(true);

// 或手动配置优化选项
auto& options = workbook->getOptions();
options.mode = WorkbookMode::STREAMING;  // 流式模式
options.compression_level = 0;           // 无压缩
options.use_shared_strings = false;      // 禁用共享字符串
```

## Workbook 类

工作簿是 Excel 文件的根容器，管理所有工作表、样式和文档属性。

### 核心方法

```cpp
namespace fastexcel::core {
class Workbook {
public:
    // === 生命周期管理 ===
    static std::unique_ptr<Workbook> create(const Path& path);
    bool open();
    bool save();
    bool saveAs(const std::string& filename);
    bool close();
    
    // === 工作表管理 ===
    std::shared_ptr<Worksheet> addWorksheet(const std::string& name = "");
    std::shared_ptr<Worksheet> insertWorksheet(size_t index, const std::string& name = "");
    bool removeWorksheet(const std::string& name);
    bool removeWorksheet(size_t index);
    std::shared_ptr<Worksheet> getWorksheet(const std::string& name);
    std::shared_ptr<Worksheet> getWorksheet(size_t index);
    
    size_t getWorksheetCount() const;
    std::vector<std::string> getWorksheetNames() const;
    bool renameWorksheet(const std::string& old_name, const std::string& new_name);
    bool moveWorksheet(size_t from_index, size_t to_index);
    void setActiveWorksheet(size_t index);
    
    // === 样式管理 ===
    int addStyle(const FormatDescriptor& style);
    int addStyle(const StyleBuilder& builder);
    std::shared_ptr<const FormatDescriptor> getStyle(int style_id) const;
    int getDefaultStyleId() const;
    bool isValidStyleId(int style_id) const;
    const FormatRepository& getStyleRepository() const;
    
    // === 文档属性 ===
    DocumentProperties& getDocumentProperties();
    const DocumentProperties& getDocumentProperties() const;
    
    // === 自定义属性 ===
    void setCustomProperty(const std::string& name, const std::string& value);
    void setCustomProperty(const std::string& name, double value);
    void setCustomProperty(const std::string& name, bool value);
    std::string getCustomProperty(const std::string& name) const;
    bool removeCustomProperty(const std::string& name);
    
    // === 定义名称 ===
    void defineName(const std::string& name, const std::string& formula, 
                   const std::string& scope = "");
    std::string getDefinedName(const std::string& name, 
                              const std::string& scope = "") const;
    bool removeDefinedName(const std::string& name, const std::string& scope = "");
    
    // === 工作簿选项 ===
    WorkbookOptions& getOptions();
    void setHighPerformanceMode(bool enable);
    void setCalcOptions(bool calc_on_load, bool full_calc_on_load);
    
    // === 工作簿保护 ===
    void protect(const std::string& password, bool lock_structure = true, 
                bool lock_windows = false);
    void unprotect();
    bool isProtected() const;
    
    // === VBA 支持 ===
    bool addVbaProject(const std::string& vba_project_path);
    bool hasVba() const;
    
    // === 统计信息 ===
    WorkbookStats getStatistics() const;
    size_t estimateMemoryUsage() const;
    size_t getTotalCellCount() const;
    
    // === 工作簿编辑功能 ===
    static std::unique_ptr<Workbook> loadForEdit(const Path& path);
    bool refresh();
    bool mergeWorkbook(const std::unique_ptr<Workbook>& other_workbook, 
                      const MergeOptions& options);
    bool exportWorksheets(const std::vector<std::string>& worksheet_names, 
                         const std::string& output_filename);
    
    // === 样式复制和传输 ===
    std::unique_ptr<StyleTransferContext> copyStylesFrom(const Workbook& source_workbook);
    FormatRepository::DeduplicationStats getStyleStats() const;
};
}
```

### 工作簿选项配置

```cpp
struct WorkbookOptions {
    WorkbookMode mode = WorkbookMode::AUTO;    // 处理模式
    int compression_level = 6;                 // ZIP压缩级别 (0-9)
    bool use_shared_strings = true;           // 启用共享字符串
    bool constant_memory = false;             // 常量内存模式
    
    // 性能调优参数
    size_t row_buffer_size = 5000;           // 行缓冲区大小
    size_t xml_buffer_size = 4 * 1024 * 1024; // XML缓冲区大小
    
    // 自动模式阈值
    size_t auto_mode_cell_threshold = 1000000;      // 100万单元格
    size_t auto_mode_memory_threshold = 100 * 1024 * 1024; // 100MB
    
    // 计算选项
    bool calc_on_load = true;                // 打开时计算
    bool full_calc_on_load = true;           // 完全重新计算
};
```

### 使用示例

```cpp
// 基本工作簿操作
auto workbook = Workbook::create("report.xlsx");
workbook->open();

// 设置文档属性
auto& props = workbook->getDocumentProperties();
props.title = "月度销售报表";
props.author = "销售部";
props.company = "我的公司";

// 添加自定义属性
workbook->setCustomProperty("部门", "销售部");
workbook->setCustomProperty("预算", 100000.0);
workbook->setCustomProperty("是否审核", true);

// 创建多个工作表
auto summary = workbook->addWorksheet("汇总");
auto detail = workbook->addWorksheet("明细");
auto chart = workbook->addWorksheet("图表");

// 设置工作表顺序
workbook->moveWorksheet(2, 1); // 将图表移动到第二位

// 启用高性能模式
workbook->setHighPerformanceMode(true);

// 保存文件
workbook->save();
```

## Worksheet 类

工作表包含单元格数据、格式设置和工作表级别的配置。

### 核心方法

```cpp
class Worksheet {
public:
    // === 基本信息 ===
    const std::string& getName() const;
    void setName(const std::string& name);
    int getSheetId() const;
    
    // === 数据写入 ===
    void writeString(int row, int col, const std::string& value, int style_id = 0);
    void writeNumber(int row, int col, double value, int style_id = 0);
    void writeBoolean(int row, int col, bool value, int style_id = 0);
    void writeDateTime(int row, int col, const std::tm& datetime, int style_id = 0);
    void writeFormula(int row, int col, const std::string& formula, int style_id = 0);
    void writeUrl(int row, int col, const std::string& url, 
                  const std::string& display_text = "", int style_id = 0);
    
    // === 数据读取 ===
    bool hasCellAt(int row, int col) const;
    const Cell& getCell(int row, int col) const;
    Cell& getCell(int row, int col);
    
    // === 范围操作 ===
    std::pair<int, int> getUsedRange() const;
    size_t getCellCount() const;
    void clear();
    void clearRange(int first_row, int first_col, int last_row, int last_col);
    
    // === 行列操作 ===
    void setColumnWidth(int col, double width);
    void setColumnWidth(int first_col, int last_col, double width);
    void setRowHeight(int row, double height);
    void setRowHeight(int first_row, int last_row, double height);
    
    void hideColumn(int col);
    void hideColumn(int first_col, int last_col);
    void hideRow(int row);
    void hideRow(int first_row, int last_row);
    
    void insertRows(int row, int count = 1);
    void insertColumns(int col, int count = 1);
    void deleteRows(int row, int count = 1);
    void deleteColumns(int col, int count = 1);
    
    // === 合并单元格 ===
    void mergeCells(int first_row, int first_col, int last_row, int last_col);
    void unmergeCells(int first_row, int first_col, int last_row, int last_col);
    bool isMergedRange(int row, int col) const;
    
    // === 自动筛选 ===
    void setAutoFilter(int first_row, int first_col, int last_row, int last_col);
    void removeAutoFilter();
    bool hasAutoFilter() const;
    
    // === 冻结窗格 ===
    void freezePanes(int row, int col);
    void freezePanes(int row, int col, int top_left_row, int top_left_col);
    void splitPanes(int row, int col);
    void removePanes();
    
    // === 工作表视图 ===
    void setZoom(int zoom_scale);
    void showGridlines(bool show = true);
    void showRowColHeaders(bool show = true);
    void setRightToLeft(bool rtl = true);
    void setTabSelected(bool selected = true);
    void setActiveCell(int row, int col);
    void setSelection(int first_row, int first_col, int last_row, int last_col);
    
    // === 打印设置 ===
    void setPrintArea(int first_row, int first_col, int last_row, int last_col);
    void setRepeatRows(int first_row, int last_row);
    void setRepeatColumns(int first_col, int last_col);
    void setPageSetup(const PageSetup& setup);
    
    // === 工作表保护 ===
    void protect(const std::string& password = "");
    void unprotect();
    bool isProtected() const;
    
    // === 数据验证 ===
    void addDataValidation(int first_row, int first_col, int last_row, int last_col,
                          const DataValidation& validation);
    
    // === 条件格式 ===
    void addConditionalFormat(int first_row, int first_col, int last_row, int last_col,
                             const ConditionalFormat& format);
    
    // === 批量搜索替换 ===
    int findAndReplace(const std::string& find_text, const std::string& replace_text,
                      bool match_case = false, bool match_entire_cell = false);
    std::vector<std::pair<int, int>> findCells(const std::string& search_text,
                                              bool match_case = false, 
                                              bool match_entire_cell = false) const;
    
    // === 内存和性能 ===
    bool isOptimizeMode() const;
    size_t getMemoryUsage() const;
    void optimizeMemory();
    
    // === XML 生成 ===
    void generateXML(const std::function<void(const char*, size_t)>& callback) const;
    void generateRelsXML(const std::function<void(const char*, size_t)>& callback) const;
};
```

### 使用示例

```cpp
// 创建工作表
auto worksheet = workbook->addWorksheet("销售数据");

// 设置列宽
worksheet->setColumnWidth(0, 15.0);  // A列
worksheet->setColumnWidth(1, 3, 12.0); // B-D列

// 写入标题行
std::vector<std::string> headers = {"产品名称", "销售日期", "数量", "金额"};
for (size_t i = 0; i < headers.size(); ++i) {
    worksheet->writeString(0, i, headers[i], header_style_id);
}

// 写入数据
worksheet->writeString(1, 0, "产品A");
worksheet->writeDateTime(1, 1, current_date, date_style_id);
worksheet->writeNumber(1, 2, 100);
worksheet->writeNumber(1, 3, 12500.0, currency_style_id);

// 合并单元格
worksheet->mergeCells(0, 0, 0, 3); // 合并标题行

// 设置自动筛选
worksheet->setAutoFilter(0, 0, 10, 3);

// 冻结窗格
worksheet->freezePanes(1, 0); // 冻结第一行

// 设置打印区域
worksheet->setPrintArea(0, 0, 10, 3);

// 查找和替换
int replaced = worksheet->findAndReplace("产品A", "Product A", false, true);

// 查找单元格
auto found_cells = worksheet->findCells("Product", false, false);
for (auto& [row, col] : found_cells) {
    std::cout << "找到匹配项: " << row << ", " << col << std::endl;
}
```

## 样式系统

FastExcel 使用现代的样式系统，包括 FormatDescriptor（样式描述符）和 StyleBuilder（样式构建器）。

### FormatDescriptor - 样式描述符

```cpp
struct FormatDescriptor {
    // 字体设置
    FontDescriptor font;
    
    // 对齐设置
    AlignmentDescriptor alignment;
    
    // 边框设置
    BorderDescriptor border;
    
    // 填充设置
    FillDescriptor fill;
    
    // 数字格式
    std::string number_format;
    
    // 保护设置
    bool locked = true;
    bool hidden = false;
    
    // 生成唯一哈希值（用于去重）
    size_t getHash() const;
    bool operator==(const FormatDescriptor& other) const;
};
```

### StyleBuilder - 样式构建器

```cpp
class StyleBuilder {
public:
    // === 字体链式设置 ===
    FontBuilder& font();
    
    // === 对齐链式设置 ===
    AlignmentBuilder& alignment();
    
    // === 边框链式设置 ===
    BorderBuilder& border();
    
    // === 填充链式设置 ===
    FillBuilder& fill();
    
    // === 数字格式 ===
    StyleBuilder& numberFormat(const std::string& format);
    
    // === 保护设置 ===
    StyleBuilder& locked(bool is_locked = true);
    StyleBuilder& hidden(bool is_hidden = true);
    
    // === 构建样式 ===
    std::shared_ptr<FormatDescriptor> build() const;
};
```

### 链式构建器详细API

```cpp
// 字体构建器
class FontBuilder {
public:
    FontBuilder& name(const std::string& font_name);
    FontBuilder& size(double font_size);
    FontBuilder& color(uint32_t color);
    FontBuilder& bold(bool is_bold = true);
    FontBuilder& italic(bool is_italic = true);
    FontBuilder& underline(UnderlineType type = UnderlineType::Single);
    FontBuilder& strikeout(bool is_strikeout = true);
};

// 对齐构建器  
class AlignmentBuilder {
public:
    AlignmentBuilder& horizontal(HorizontalAlign align);
    AlignmentBuilder& vertical(VerticalAlign align);
    AlignmentBuilder& wrap(bool text_wrap = true);
    AlignmentBuilder& rotation(int angle);
    AlignmentBuilder& indent(int level);
    AlignmentBuilder& shrinkToFit(bool shrink = true);
};

// 边框构建器
class BorderBuilder {
public:
    BorderBuilder& all(BorderStyle style, uint32_t color = 0x000000);
    BorderBuilder& left(BorderStyle style, uint32_t color = 0x000000);
    BorderBuilder& right(BorderStyle style, uint32_t color = 0x000000);
    BorderBuilder& top(BorderStyle style, uint32_t color = 0x000000);
    BorderBuilder& bottom(BorderStyle style, uint32_t color = 0x000000);
    BorderBuilder& diagonal(BorderStyle style, uint32_t color = 0x000000);
};

// 填充构建器
class FillBuilder {
public:
    FillBuilder& backgroundColor(uint32_t color);
    FillBuilder& foregroundColor(uint32_t color);  
    FillBuilder& pattern(PatternType pattern);
};
```

### 样式使用示例

```cpp
// 使用 StyleBuilder 创建样式
StyleBuilder title_builder;
title_builder.font()
    .name("Arial")
    .size(16)
    .bold(true)
    .color(0xFFFFFF);
title_builder.fill()
    .backgroundColor(0x4472C4);
title_builder.alignment()
    .horizontal(HorizontalAlign::Center)
    .vertical(VerticalAlign::Center);
title_builder.border()
    .all(BorderStyle::Thin);

int title_style_id = workbook->addStyle(title_builder);

// 创建货币样式
StyleBuilder currency_builder;
currency_builder.numberFormat("¥#,##0.00");
currency_builder.alignment().horizontal(HorizontalAlign::Right);
int currency_style_id = workbook->addStyle(currency_builder);

// 创建日期样式
StyleBuilder date_builder;
date_builder.numberFormat("yyyy-mm-dd");
int date_style_id = workbook->addStyle(date_builder);

// 应用样式
worksheet->writeString(0, 0, "销售报表", title_style_id);
worksheet->writeNumber(1, 0, 12345.67, currency_style_id);
worksheet->writeDateTime(2, 0, current_time, date_style_id);

// 检查样式统计
auto stats = workbook->getStyleStats();
std::cout << "创建样式: " << stats.created_count << std::endl;
std::cout << "去重样式: " << stats.deduplicated_count << std::endl;
```

### 预定义样式常量

```cpp
namespace fastexcel::core {
// 颜色常量
constexpr uint32_t COLOR_BLACK   = 0x000000;
constexpr uint32_t COLOR_WHITE   = 0xFFFFFF;
constexpr uint32_t COLOR_RED     = 0xFF0000;
constexpr uint32_t COLOR_GREEN   = 0x00FF00;
constexpr uint32_t COLOR_BLUE    = 0x0000FF;
constexpr uint32_t COLOR_YELLOW  = 0xFFFF00;

// Excel主题色
constexpr uint32_t THEME_BLUE    = 0x4472C4;
constexpr uint32_t THEME_RED     = 0xC5504B;
constexpr uint32_t THEME_GREEN   = 0x70AD47;
constexpr uint32_t THEME_ORANGE  = 0xE07C24;

// 常用数字格式
constexpr const char* NUMBER_FORMAT_GENERAL    = "General";
constexpr const char* NUMBER_FORMAT_INTEGER    = "0";
constexpr const char* NUMBER_FORMAT_DECIMAL    = "0.00";
constexpr const char* NUMBER_FORMAT_CURRENCY   = "¥#,##0.00";
constexpr const char* NUMBER_FORMAT_PERCENTAGE = "0.00%";
constexpr const char* NUMBER_FORMAT_DATE       = "yyyy-mm-dd";
constexpr const char* NUMBER_FORMAT_TIME       = "h:mm:ss";
constexpr const char* NUMBER_FORMAT_DATETIME   = "yyyy-mm-dd h:mm:ss";
}
```

## Cell 类

单元格类存储单元格数据和相关信息。

```cpp
enum class CellType {
    Empty, String, Number, Boolean, Date, Formula, Error
};

class Cell {
public:
    // === 构造和析构 ===
    Cell() = default;
    Cell(const Cell& other);
    Cell& operator=(const Cell& other);
    Cell(Cell&& other) noexcept = default;
    Cell& operator=(Cell&& other) noexcept = default;
    
    // === 设置值 ===
    void setValue(const std::string& value);
    void setValue(double value);
    void setValue(bool value);
    void setValue(int value);
    void setFormula(const std::string& formula);
    
    // === 获取值 ===
    CellType getType() const;
    std::string getStringValue() const;
    double getNumberValue() const;
    bool getBooleanValue() const;
    std::string getFormula() const;
    
    // === 样式设置 ===
    void setStyleId(int style_id);
    int getStyleId() const;
    
    // === 超链接 ===
    void setHyperlink(const std::string& url);
    std::string getHyperlink() const;
    bool hasHyperlink() const;
    
    // === 检查状态 ===
    bool isEmpty() const;
    bool isString() const;
    bool isNumber() const;
    bool isBoolean() const;
    bool isFormula() const;
    
    // === 内存管理 ===
    size_t getMemoryUsage() const;
    void clear();
};
```

### 使用示例

```cpp
// 获取单元格引用
auto& cell = worksheet->getCell(0, 0);

// 设置不同类型的值
cell.setValue("Hello World");
cell.setValue(123.45);
cell.setValue(true);
cell.setFormula("SUM(A1:A10)");

// 应用样式
cell.setStyleId(header_style_id);

// 添加超链接
cell.setHyperlink("https://www.example.com");

// 检查单元格类型
if (cell.isString()) {
    std::cout << "字符串值: " << cell.getStringValue() << std::endl;
}
if (cell.isNumber()) {
    std::cout << "数值: " << cell.getNumberValue() << std::endl;
}
if (cell.hasHyperlink()) {
    std::cout << "超链接: " << cell.getHyperlink() << std::endl;
}
```

## 流式处理

FastExcel 支持三种流式处理模式，可以处理不同规模的数据。

### WorkbookMode 模式选择

```cpp
enum class WorkbookMode {
    AUTO,      // 自动选择（推荐）
    BATCH,     // 批量模式：内存缓存后批量写入
    STREAMING  // 流式模式：实时写入，恒定内存
};

// 设置处理模式
auto& options = workbook->getOptions();
options.mode = WorkbookMode::STREAMING;

// 或者使用高性能模式（自动优化）
workbook->setHighPerformanceMode(true);
```

### 性能模式对比

| 模式 | 内存使用 | 处理速度 | 适用场景 |
|------|----------|----------|----------|
| **BATCH** | 高（全量缓存） | 快 | 小到中型文件 (<50MB) |
| **STREAMING** | 恒定（8KB缓冲） | 中等 | 大文件，内存受限 (>100MB) |
| **AUTO** | 智能选择 | 最优 | 所有场景（推荐） |

### 高性能配置示例

```cpp
// 极致性能配置
workbook->setHighPerformanceMode(true);

// 或者手动配置
auto& options = workbook->getOptions();
options.mode = WorkbookMode::AUTO;
options.compression_level = 0;           // 无压缩
options.use_shared_strings = false;      // 禁用共享字符串
options.constant_memory = true;          // 启用恒定内存
options.row_buffer_size = 10000;         // 大缓冲区
options.xml_buffer_size = 8 * 1024 * 1024; // 8MB XML缓冲

// 设置自动模式阈值
options.auto_mode_cell_threshold = 2000000;    // 200万单元格
options.auto_mode_memory_threshold = 200 * 1024 * 1024; // 200MB
```

## 文档属性管理

### DocumentProperties 结构

```cpp
struct DocumentProperties {
    std::string title;       // 文档标题
    std::string author;      // 作者
    std::string manager;     // 管理者
    std::string company;     // 公司
    std::string category;    // 类别
    std::string keywords;    // 关键词
    std::string comments;    // 备注
    std::string subject;     // 主题
    std::string status;      // 状态
    
    std::tm created_time;    // 创建时间
    std::tm modified_time;   // 修改时间
};
```

### CustomPropertyManager

```cpp
class CustomPropertyManager {
public:
    // 设置属性
    void setProperty(const std::string& name, const std::string& value);
    void setProperty(const std::string& name, double value);
    void setProperty(const std::string& name, bool value);
    
    // 获取属性
    std::string getProperty(const std::string& name) const;
    bool hasProperty(const std::string& name) const;
    bool removeProperty(const std::string& name);
    
    // 批量操作
    std::unordered_map<std::string, std::string> getAllProperties() const;
    void clear();
    bool empty() const;
    size_t size() const;
};
```

### 使用示例

```cpp
// 设置文档属性
auto& props = workbook->getDocumentProperties();
props.title = "2024年销售报表";
props.author = "销售部";
props.company = "我的公司";
props.category = "财务报表";
props.keywords = "销售,报表,2024";
props.comments = "月度销售数据汇总";
props.status = "草稿";

// 设置自定义属性
workbook->setCustomProperty("部门", "销售部");
workbook->setCustomProperty("版本", 1.2);
workbook->setCustomProperty("已审核", false);
workbook->setCustomProperty("截止日期", "2024-12-31");

// 获取自定义属性
std::string dept = workbook->getCustomProperty("部门");
bool reviewed = workbook->getCustomProperty("已审核") == "true";

// 移除属性
workbook->removeCustomProperty("临时属性");
```

## 工具类

### 单元格引用工具

```cpp
namespace fastexcel::utils {
// 单元格引用转换
std::string cellReference(int row, int col);           // (0,0) -> "A1"
std::pair<int, int> parseReference(const std::string& ref); // "A1" -> (0,0)

// 列转换
std::string columnToLetter(int col);    // 0 -> "A", 25 -> "Z", 26 -> "AA"
int letterToColumn(const std::string& letter); // "A" -> 0, "AA" -> 26

// 范围引用
std::string rangeReference(int first_row, int first_col, 
                          int last_row, int last_col); // "A1:B2"
}
```

### 颜色工具

```cpp
namespace fastexcel::utils {
// RGB转换
uint32_t rgbToColor(int r, int g, int b);
std::tuple<int, int, int> colorToRgb(uint32_t color);

// 颜色调整
uint32_t darkenColor(uint32_t color, double factor);  // 变暗
uint32_t lightenColor(uint32_t color, double factor); // 变亮
uint32_t blendColors(uint32_t color1, uint32_t color2, double ratio);
}
```

### 日期时间工具

```cpp
namespace fastexcel::utils {
// Excel日期序列号转换
double dateTimeToSerial(const std::tm& datetime);
std::tm serialToDateTime(double serial);

// 当前时间
std::tm getCurrentTime();
std::string formatTimeISO8601(const std::tm& time);
}
```

### 使用示例

```cpp
// 单元格引用
std::string ref = fastexcel::utils::cellReference(0, 0); // "A1"
auto [row, col] = fastexcel::utils::parseReference("B2"); // (1, 1)

// 范围引用
std::string range = fastexcel::utils::rangeReference(0, 0, 9, 3); // "A1:D10"

// 颜色操作
uint32_t red = fastexcel::utils::rgbToColor(255, 0, 0);
auto [r, g, b] = fastexcel::utils::colorToRgb(0xFF0000);
uint32_t dark_red = fastexcel::utils::darkenColor(red, 0.8);

// 日期时间
auto now = fastexcel::utils::getCurrentTime();
double excel_date = fastexcel::utils::dateTimeToSerial(now);
```

## 最佳实践

### 1. 性能优化指南

```cpp
// ✅ 推荐：使用自动模式
auto& options = workbook->getOptions();
options.mode = WorkbookMode::AUTO;

// ✅ 推荐：重用样式
StyleBuilder header_style;
header_style.font().bold(true).color(0xFFFFFF);
int header_id = workbook->addStyle(header_style);

// 在循环中重用样式ID
for (int col = 0; col < 10; ++col) {
    worksheet->writeString(0, col, headers[col], header_id);
}

// ❌ 避免：在循环中创建样式
for (int col = 0; col < 10; ++col) {
    StyleBuilder style;  // 每次都创建新样式
    style.font().bold(true);
    int style_id = workbook->addStyle(style);
    worksheet->writeString(0, col, headers[col], style_id);
}
```

### 2. 内存管理

```cpp
// ✅ 大数据处理：启用流式模式
workbook->getOptions().mode = WorkbookMode::STREAMING;

// ✅ 分批处理
const size_t BATCH_SIZE = 5000;
for (size_t i = 0; i < total_rows; i += BATCH_SIZE) {
    size_t end = std::min(i + BATCH_SIZE, total_rows);
    processBatch(worksheet, i, end);
    
    // 可选：报告进度
    if (i % (BATCH_SIZE * 10) == 0) {
        std::cout << "已处理: " << i << " / " << total_rows << std::endl;
    }
}
```

### 3. 错误处理

```cpp
try {
    auto workbook = Workbook::create("report.xlsx");
    workbook->open();
    
    auto worksheet = workbook->addWorksheet("数据");
    // ... 数据操作
    
    workbook->save();
    
} catch (const std::exception& e) {
    LOG_ERROR("创建Excel文件失败: {}", e.what());
    // 错误恢复逻辑
    return false;
}
```

### 4. 资源管理

```cpp
// ✅ RAII模式 - 自动管理资源
{
    auto workbook = Workbook::create("temp.xlsx");
    workbook->open();
    // ... 操作
    workbook->save();
} // 自动释放资源

// ✅ 智能指针管理
std::shared_ptr<Worksheet> sheet = workbook->addWorksheet("Sheet1");
// worksheet会在workbook析构时自动释放
```

### 5. 样式复制和迁移

```cpp
// 从现有工作簿复制样式
auto source_workbook = Workbook::loadForEdit("template.xlsx");
auto transfer_context = workbook->copyStylesFrom(*source_workbook);

// 使用ID映射应用样式
int source_style_id = 5;
int target_style_id = transfer_context->mapStyleId(source_style_id);
worksheet->writeString(0, 0, "Hello", target_style_id);

// 查看传输统计
auto stats = transfer_context->getTransferStats();
LOG_INFO("传输了{}个样式，去重{}个", stats.transferred_count, stats.deduplicated_count);
```

### 6. 多工作表操作

```cpp
// 批量重命名工作表
std::unordered_map<std::string, std::string> rename_map = {
    {"Sheet1", "汇总"},
    {"Sheet2", "明细"},
    {"Sheet3", "图表"}
};
int renamed = workbook->batchRenameWorksheets(rename_map);

// 批量删除工作表
std::vector<std::string> to_remove = {"临时表", "测试表"};
int removed = workbook->batchRemoveWorksheets(to_remove);

// 重新排列工作表顺序
std::vector<std::string> new_order = {"汇总", "图表", "明细"};
workbook->reorderWorksheets(new_order);
```

### 7. 查找和替换

```cpp
// 全局查找替换
FindReplaceOptions options;
options.match_case = false;
options.match_entire_cell = false;
options.worksheet_filter = {"Sheet1", "Sheet2"}; // 仅在指定工作表中查找

int replaced = workbook->findAndReplaceAll("旧文本", "新文本", options);

// 查找所有匹配项
auto results = workbook->findAll("搜索文本", options);
for (auto& [sheet_name, row, col] : results) {
    std::cout << "找到: " << sheet_name << " " << row << ":" << col << std::endl;
}
```

### 8. 完整的应用示例

```cpp
#include "fastexcel/FastExcel.hpp"
using namespace fastexcel::core;

void createSalesReport() {
    try {
        // 创建工作簿
        auto workbook = Workbook::create("销售报表2024.xlsx");
        workbook->open();
        
        // 启用高性能模式
        workbook->setHighPerformanceMode(true);
        
        // 设置文档属性
        auto& props = workbook->getDocumentProperties();
        props.title = "2024年销售报表";
        props.author = "销售部";
        props.company = "我的公司";
        
        // 创建样式
        StyleBuilder title_style;
        title_style.font().size(18).bold(true).color(0xFFFFFF);
        title_style.fill().backgroundColor(0x4472C4);
        title_style.alignment().horizontal(HorizontalAlign::Center);
        int title_id = workbook->addStyle(title_style);
        
        StyleBuilder header_style;
        header_style.font().bold(true);
        header_style.fill().backgroundColor(0xD9E1F2);
        header_style.border().all(BorderStyle::Thin);
        int header_id = workbook->addStyle(header_style);
        
        StyleBuilder currency_style;
        currency_style.numberFormat("¥#,##0.00");
        currency_style.border().all(BorderStyle::Thin);
        int currency_id = workbook->addStyle(currency_style);
        
        // 创建工作表
        auto worksheet = workbook->addWorksheet("销售数据");
        
        // 设置列宽
        worksheet->setColumnWidth(0, 15);
        worksheet->setColumnWidth(1, 12);
        worksheet->setColumnWidth(2, 10);
        worksheet->setColumnWidth(3, 12);
        
        // 写入标题
        worksheet->mergeCells(0, 0, 0, 3);
        worksheet->writeString(0, 0, "2024年销售报表", title_id);
        
        // 写入表头
        std::vector<std::string> headers = {"产品名称", "销售日期", "数量", "金额"};
        for (size_t i = 0; i < headers.size(); ++i) {
            worksheet->writeString(1, i, headers[i], header_id);
        }
        
        // 写入数据（模拟大量数据）
        for (int row = 2; row < 10002; ++row) {
            worksheet->writeString(row, 0, "产品" + std::to_string(row - 1));
            worksheet->writeString(row, 1, "2024-01-01");
            worksheet->writeNumber(row, 2, row * 10);
            worksheet->writeNumber(row, 3, row * 123.45, currency_id);
            
            // 进度报告
            if (row % 1000 == 0) {
                std::cout << "已写入: " << row << " 行\n";
            }
        }
        
        // 设置自动筛选
        worksheet->setAutoFilter(1, 0, 10001, 3);
        
        // 冻结窗格
        worksheet->freezePanes(2, 0);
        
        // 保存文件
        auto start_time = std::chrono::high_resolution_clock::now();
        workbook->save();
        auto end_time = std::chrono::high_resolution_clock::now();
        
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time);
        std::cout << "保存耗时: " << duration.count() << " 毫秒\n";
        
        // 查看统计信息
        auto stats = workbook->getStatistics();
        std::cout << "总工作表: " << stats.total_worksheets << "\n";
        std::cout << "总单元格: " << stats.total_cells << "\n";
        std::cout << "总样式: " << stats.total_formats << "\n";
        std::cout << "内存使用: " << stats.memory_usage / 1024 / 1024 << " MB\n";
        
        auto style_stats = workbook->getStyleStats();
        std::cout << "样式去重: " << style_stats.deduplicated_count << "\n";
        
    } catch (const std::exception& e) {
        std::cerr << "创建报表失败: " << e.what() << std::endl;
    }
}
```

---

**FastExcel** - 现代化、高性能的 C++ Excel 库

*本文档版本: 2.0.0*  
*基于代码版本: 当前最新*  
*最后更新: 2025-01-06*

如有问题或建议，请提交 [Issue](https://github.com/your-repo/FastExcel/issues)