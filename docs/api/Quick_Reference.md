# FastExcel API 快速参考

本文档提供FastExcel最常用API的快速参考，基于当前源码结构(v2.0)。

## 🚀 基本使用

### 1. 创建和保存Excel文件

```cpp
#include "fastexcel/FastExcel.hpp"
using namespace fastexcel;
using namespace fastexcel::core;

// 创建新文件
auto workbook = std::make_unique<Workbook>(Path("output.xlsx"));

// 初始化工作簿
if (!workbook->open()) {
    throw std::runtime_error("无法创建工作簿");
}

// 添加工作表
auto worksheet = workbook->addWorksheet("数据表");

// 保存文件
workbook->save();
```

### 2. 单元格操作

```cpp
// 获取单元格
const auto& cell = worksheet->getCell(0, 0);          // 使用行列索引 (0-based)
bool hasCell = worksheet->hasCellAt(0, 0);           // 检查单元格是否存在

// 设置值
worksheet->setValue(0, 0, "Hello World");            // 字符串
worksheet->setValue(0, 1, 123.45);                   // 数字
worksheet->setValue(0, 2, true);                     // 布尔值

// 设置公式
worksheet->setFormula(1, 0, "=SUM(A1:C1)", 0.0);    // 公式和结果缓存

// 获取值（需要类型检查）
const auto& cell = worksheet->getCell(0, 0);
if (cell.isString()) {
    std::string value = cell.getStringValue();
} else if (cell.isNumber()) {
    double value = cell.getNumberValue();
} else if (cell.isBoolean()) {
    bool value = cell.getBooleanValue();
}
```

### 3. 样式设置

```cpp
// 创建样式构建器
auto styleBuilder = workbook->createStyleBuilder();

// 构建样式
auto headerStyle = styleBuilder
    .fontName("微软雅黑")
    .fontSize(14)
    .bold(true)
    .fontColor(Color::WHITE)
    .backgroundColor(Color::DARK_BLUE)
    .horizontalAlign(HorizontalAlign::Center)
    .build();

// 添加样式到仓储
int styleIndex = workbook->addFormat(headerStyle);

// 应用样式到单元格
auto formatDescriptor = workbook->getFormatDescriptor(styleIndex);
worksheet->getCell(0, 0).setFormat(formatDescriptor);
```

## 📊 工作簿管理

### Workbook 类核心方法

```cpp
class Workbook {
public:
    // 构造函数
    explicit Workbook(const Path& filename);
    
    // 初始化
    bool open();
    
    // 工作表管理
    std::shared_ptr<Worksheet> addWorksheet(const std::string& name);
    std::shared_ptr<Worksheet> getWorksheet(const std::string& name);
    std::shared_ptr<Worksheet> getWorksheet(size_t index);
    std::vector<std::string> getSheetNames() const;
    size_t getWorksheetCount() const;
    
    // 样式管理
    StyleBuilder createStyleBuilder();
    int addFormat(const FormatDescriptor& format);
    std::shared_ptr<const FormatDescriptor> getFormatDescriptor(int formatIndex) const;
    FormatRepository& getStyles();
    
    // 共享字符串
    SharedStringTable* getSharedStrings();
    
    // 保存操作
    bool save();
    bool saveAs(const Path& new_path);
    
    // 选项设置
    WorkbookOptions& getOptions();
    void setOptions(const WorkbookOptions& options);
};
```

### 工作簿选项

```cpp
struct WorkbookOptions {
    bool constant_memory = false;           // 常量内存模式
    WorkbookMode mode = WorkbookMode::AUTO; // 工作簿模式
    bool use_shared_strings = true;        // 使用共享字符串表
    bool use_zip64 = false;                // 使用ZIP64格式
    bool optimize_for_speed = false;       // 优化速度
    size_t row_buffer_size = 5000;         // 行缓冲区大小
    int compression_level = 6;             // 压缩级别 (0-9)
    size_t xml_buffer_size = 4 * 1024 * 1024; // XML缓冲区大小
};
```

## 📄 工作表操作

### Worksheet 类核心方法

```cpp
class Worksheet {
public:
    // 单元格访问
    Cell& getCell(int row, int col);
    const Cell& getCell(int row, int col) const;
    bool hasCellAt(int row, int col) const;
    
    // 数据设置
    void setValue(int row, int col, const std::string& value);
    void setValue(int row, int col, double value);
    void setValue(int row, int col, bool value);
    void setFormula(int row, int col, const std::string& formula, double result);
    
    // 行列操作
    void setColumnWidth(int col, double width);
    void setRowHeight(int row, double height);
    double getColumnWidth(int col) const;
    double getRowHeight(int row) const;
    
    // 合并单元格
    void mergeCells(int startRow, int startCol, int endRow, int endCol);
    std::vector<CellRange> getMergeRanges() const;
    
    // 自动筛选
    void setAutoFilter(const CellRange& range);
    bool hasAutoFilter() const;
    
    // 冻结窗格
    void freezePanes(int row, int col);
    bool hasFrozenPanes() const;
    
    // 工作表信息
    std::string getName() const;
    void setName(const std::string& name);
    std::pair<int, int> getUsedRange() const;
    size_t getCellCount() const;
    
    // 工作表保护
    void protectSheet(const std::string& password);
    bool isProtected() const;
};
```

## 🎨 样式系统

### StyleBuilder 链式API

```cpp
class StyleBuilder {
public:
    // 字体设置
    StyleBuilder& fontName(const std::string& name);
    StyleBuilder& fontSize(double size);
    StyleBuilder& bold(bool is_bold = true);
    StyleBuilder& italic(bool is_italic = true);
    StyleBuilder& fontColor(const Color& color);
    
    // 填充设置
    StyleBuilder& backgroundColor(const Color& color);
    StyleBuilder& pattern(PatternType pattern);
    
    // 边框设置
    StyleBuilder& leftBorder(BorderStyle style, const Color& color);
    StyleBuilder& rightBorder(BorderStyle style, const Color& color);
    StyleBuilder& topBorder(BorderStyle style, const Color& color);
    StyleBuilder& bottomBorder(BorderStyle style, const Color& color);
    StyleBuilder& allBorders(BorderStyle style, const Color& color);
    
    // 对齐设置
    StyleBuilder& horizontalAlign(HorizontalAlign align);
    StyleBuilder& verticalAlign(VerticalAlign align);
    StyleBuilder& textWrap(bool wrap = true);
    
    // 数字格式
    StyleBuilder& numberFormat(const std::string& format);
    
    // 构建样式
    FormatDescriptor build() const;
};
```

### 颜色系统

```cpp
class Color {
public:
    // 预定义颜色
    static const Color BLACK;
    static const Color WHITE;
    static const Color RED;
    static const Color GREEN;
    static const Color BLUE;
    static const Color YELLOW;
    static const Color DARK_BLUE;
    
    // 创建颜色
    static Color fromRGB(uint8_t r, uint8_t g, uint8_t b);
    static Color fromHex(const std::string& hex);
    static Color fromTheme(int theme_index, double tint = 0.0);
    
    // 颜色信息
    uint8_t getRed() const;
    uint8_t getGreen() const;
    uint8_t getBlue() const;
    std::string toHex() const;
};
```

## 🔧 实用工具

### 地址和范围

```cpp
// 单元格范围
struct CellRange {
    int first_row;
    int first_col;
    int last_row;
    int last_col;
    
    CellRange(int fr, int fc, int lr, int lc);
    bool contains(int row, int col) const;
    size_t getCellCount() const;
};

// 地址解析工具
namespace utils {
    class AddressParser {
    public:
        static std::pair<int, int> parseA1(const std::string& address);
        static std::string formatA1(int row, int col);
        static CellRange parseRange(const std::string& range);
    };
}
```

### 性能监控

```cpp
// 工作簿统计
struct WorkbookStats {
    size_t total_worksheets;
    size_t total_cells;
    size_t total_formats;
    size_t memory_usage;
    std::unordered_map<std::string, size_t> worksheet_cell_counts;
};

// 获取统计信息
WorkbookStats stats = workbook->getStatistics();
```

## 🏃‍♂️ 高性能操作

### 工作簿模式选择

```cpp
enum class WorkbookMode {
    AUTO = 0,       // 自动选择模式
    BATCH = 1,      // 批量处理模式
    STREAMING = 2   // 流式处理模式
};

// 设置模式
WorkbookOptions options;
options.mode = WorkbookMode::BATCH;        // 大数据处理
options.constant_memory = true;           // 常量内存使用
workbook->setOptions(options);
```

### 样式优化

```cpp
// 样式去重统计
struct DeduplicationStats {
    size_t total_requests;
    size_t unique_formats;
    double deduplication_ratio;
};

// 获取去重统计
auto& formatRepo = workbook->getStyles();
DeduplicationStats stats = formatRepo.getDeduplicationStats();
```

## 📝 最佳实践

### 1. 工作簿生命周期管理

```cpp
// ✅ 正确的创建方式
auto workbook = std::make_unique<Workbook>(Path("output.xlsx"));
if (!workbook->open()) {
    throw std::runtime_error("无法创建工作簿");
}

// 使用workbook...

// ❌ 错误：忘记调用open()
auto workbook = std::make_unique<Workbook>(Path("output.xlsx"));
// 直接使用会导致未初始化错误
```

### 2. 样式复用优化

```cpp
// ✅ 样式复用
auto commonStyle = workbook->createStyleBuilder()
    .fontName("Calibri")
    .fontSize(11)
    .build();
int styleIndex = workbook->addFormat(commonStyle);

// 在多处使用相同样式
for (int i = 0; i < 100; ++i) {
    auto formatDescriptor = workbook->getFormatDescriptor(styleIndex);
    worksheet->getCell(i, 0).setFormat(formatDescriptor);
}
```

### 3. 错误处理

```cpp
try {
    auto workbook = std::make_unique<Workbook>(Path("test.xlsx"));
    if (!workbook->open()) {
        throw std::runtime_error("无法创建工作簿");
    }
    
    auto worksheet = workbook->addWorksheet("Test");
    worksheet->setValue(0, 0, "Hello World");
    
    if (!workbook->save()) {
        throw std::runtime_error("保存失败");
    }
} catch (const std::exception& e) {
    std::cerr << "错误: " << e.what() << std::endl;
}
```

## 🔍 调试和诊断

### 单元格类型检查

```cpp
const auto& cell = worksheet->getCell(0, 0);

// 类型检查
if (cell.isEmpty()) {
    std::cout << "单元格为空" << std::endl;
} else if (cell.isString()) {
    std::cout << "字符串值: " << cell.getStringValue() << std::endl;
} else if (cell.isNumber()) {
    std::cout << "数值: " << cell.getNumberValue() << std::endl;
} else if (cell.isBoolean()) {
    std::cout << "布尔值: " << cell.getBooleanValue() << std::endl;
} else if (cell.isFormula()) {
    std::cout << "公式: " << cell.getFormula() << std::endl;
}
```

### 内存使用监控

```cpp
// 获取内存使用统计
WorkbookStats stats = workbook->getStatistics();
std::cout << "总内存使用: " << stats.memory_usage << " 字节" << std::endl;
std::cout << "总单元格数: " << stats.total_cells << std::endl;
std::cout << "样式数量: " << stats.total_formats << std::endl;
```

---

## 📚 相关文档

- [完整API参考](core-api.md) - 详细的类和方法文档
- [架构概览](../architecture/overview.md) - 系统设计说明
- [示例教程](../examples-tutorial.md) - 完整使用示例

---

*FastExcel API 快速参考 v2.0 - 基于当前源码结构更新*
*最后更新: 2025-08-24*