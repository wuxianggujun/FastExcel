# FastExcel 核心 API 参考 v2.0

本文档提供 FastExcel 核心API的详细参考信息，基于当前源码结构(v2.0)。

## 📋 目录

- [Workbook 工作簿](#workbook-工作簿)
- [Worksheet 工作表](#worksheet-工作表) 
- [Cell 单元格](#cell-单元格)
- [StyleBuilder 样式构建器](#stylebuilder-样式构建器)
- [FormatDescriptor 格式描述符](#formatdescriptor-格式描述符)
- [Color 颜色系统](#color-颜色系统)

---

## Workbook 工作簿

### 创建和初始化

```cpp
class Workbook {
public:
    // 构造函数
    explicit Workbook(const Path& filename);
    
    // 初始化工作簿（必须调用）
    bool open();
    
    // 使用示例
    auto workbook = std::make_unique<Workbook>(Path("output.xlsx"));
    if (!workbook->open()) {
        throw std::runtime_error("无法创建工作簿");
    }
};
```

### 工作表管理

```cpp
// 添加工作表
std::shared_ptr<Worksheet> addWorksheet(const std::string& name);

// 获取工作表
std::shared_ptr<Worksheet> getWorksheet(size_t index);
std::shared_ptr<Worksheet> getWorksheet(const std::string& name);

// 工作表信息
size_t getWorksheetCount() const;
std::vector<std::string> getSheetNames() const;
```

### 样式管理

```cpp
// 创建样式构建器
StyleBuilder createStyleBuilder();

// 格式管理
int addFormat(const FormatDescriptor& format);
std::shared_ptr<const FormatDescriptor> getFormatDescriptor(int formatIndex) const;
FormatRepository& getStyles();

// 共享字符串管理
SharedStringTable* getSharedStrings();
```

### 保存和选项

```cpp
// 保存文件
bool save();
bool saveAs(const Path& new_path);

// 工作簿选项
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

WorkbookOptions& getOptions();
void setOptions(const WorkbookOptions& options);

enum class WorkbookMode {
    AUTO = 0,       // 自动选择模式
    BATCH = 1,      // 批量处理模式
    STREAMING = 2   // 流式处理模式
};
```

### 统计和诊断

```cpp
// 性能统计
struct WorkbookStats {
    size_t total_worksheets;
    size_t total_cells;
    size_t total_formats;
    size_t memory_usage;
    std::unordered_map<std::string, size_t> worksheet_cell_counts;
};
WorkbookStats getStatistics() const;

// 内存使用
size_t getTotalMemoryUsage() const;
bool isModified() const;
```

---

## Worksheet 工作表

### 基本信息

```cpp
class Worksheet {
public:
    // 工作表属性
    std::string getName() const;
    void setName(const std::string& name);
    int getSheetId() const;
    
    // 数据范围
    std::pair<int, int> getUsedRange() const;
    size_t getCellCount() const;
    
    // 状态检查
    bool isEmpty() const;
};
```

### 单元格操作

```cpp
// 获取单元格（0-based索引）
Cell& getCell(int row, int col);
const Cell& getCell(int row, int col) const;
bool hasCellAt(int row, int col) const;

// 设置数据
void setValue(int row, int col, const std::string& value);
void setValue(int row, int col, double value);
void setValue(int row, int col, bool value);
void setFormula(int row, int col, const std::string& formula, double result = 0.0);
```

### 行列操作

```cpp
// 行操作
void setRowHeight(int row, double height);
double getRowHeight(int row) const;
void hideRow(int row, bool hidden = true);
bool isRowHidden(int row) const;

// 列操作  
void setColumnWidth(int col, double width);
double getColumnWidth(int col) const;
void hideColumn(int col, bool hidden = true);
bool isColumnHidden(int col) const;

// 列格式设置
void setColumnFormat(int col, std::shared_ptr<const FormatDescriptor> format);
std::shared_ptr<const FormatDescriptor> getColumnFormat(int col) const;

// 行格式设置
void setRowFormat(int row, std::shared_ptr<const FormatDescriptor> format);
std::shared_ptr<const FormatDescriptor> getRowFormat(int row) const;
```

### 合并单元格

```cpp
// 合并操作
void mergeCells(int start_row, int start_col, int end_row, int end_col);
std::vector<CellRange> getMergeRanges() const;

// 合并信息
bool isCellMerged(int row, int col) const;

// 范围结构
struct CellRange {
    int first_row;
    int first_col;
    int last_row;
    int last_col;
};
```

### 高级功能

```cpp
// 自动筛选
void setAutoFilter(const CellRange& range);
void clearAutoFilter();
bool hasAutoFilter() const;
CellRange getAutoFilterRange() const;

// 冻结窗格
void freezePanes(int row, int col);
bool hasFrozenPanes() const;

// 工作表保护
void protectSheet(const std::string& password);
bool isProtected() const;
std::string getProtectionPassword() const;

// 图片插入
void insertImage(int row, int col, const std::string& imagePath);
bool hasImages() const;
```

---

## Cell 单元格

### 数据类型

```cpp
enum class CellType : uint8_t {
    Empty = 0,
    Number = 1,
    String = 2,
    Boolean = 3,
    Formula = 4,
    Date = 5,
    Error = 6,
    Hyperlink = 7
};
```

### 基本操作

```cpp
class Cell {
public:
    // 类型检查
    CellType getType() const;
    bool isEmpty() const;
    bool isNumber() const;
    bool isString() const;
    bool isBoolean() const;
    bool isFormula() const;
    
    // 值访问（类型安全）
    std::string getStringValue() const;   // 仅当 isString() 为 true
    double getNumberValue() const;        // 仅当 isNumber() 为 true  
    bool getBooleanValue() const;         // 仅当 isBoolean() 为 true
    
    // 模板化访问（新增，但需谨慎使用）
    template<typename T>
    T getValue() const;
};
```

### 公式支持

```cpp
// 公式操作
void setFormula(const std::string& formula, double result = 0.0);
std::string getFormula() const;
double getFormulaResult() const;

// 共享公式
void setSharedFormula(int shared_index, double result = 0.0);
int getSharedFormulaIndex() const;
bool isSharedFormula() const;
```

### 格式和样式

```cpp
// 格式操作
void setFormat(std::shared_ptr<const FormatDescriptor> format);
std::shared_ptr<const FormatDescriptor> getFormatDescriptor() const;
bool hasFormat() const;

// 超链接
void setHyperlink(const std::string& url);
std::string getHyperlink() const;
bool hasHyperlink() const;
```

---

## StyleBuilder 样式构建器

### 字体设置

```cpp
class StyleBuilder {
public:
    // 字体基础属性
    StyleBuilder& fontName(const std::string& name);
    StyleBuilder& fontSize(double size);
    StyleBuilder& fontColor(const Color& color);
    
    // 字体样式
    StyleBuilder& bold(bool is_bold = true);
    StyleBuilder& italic(bool is_italic = true);
    StyleBuilder& underline(UnderlineType type = UnderlineType::Single);
    StyleBuilder& strikeout(bool is_strikeout = true);
    
    // 组合设置
    StyleBuilder& font(const std::string& name, double size);
    StyleBuilder& font(const std::string& name, double size, bool is_bold);
};
```

### 对齐设置

```cpp
// 对齐方式
StyleBuilder& horizontalAlign(HorizontalAlign align);
StyleBuilder& verticalAlign(VerticalAlign align);
StyleBuilder& textWrap(bool wrap = true);
StyleBuilder& rotation(int16_t degrees);
StyleBuilder& indent(uint8_t level);
StyleBuilder& shrinkToFit(bool shrink = true);

// 对齐枚举
enum class HorizontalAlign {
    None, Left, Center, Right, Fill, Justify, 
    CenterContinuous, Distributed
};

enum class VerticalAlign {
    Top, Center, Bottom, Justify, Distributed
};
```

### 边框设置

```cpp
// 边框样式
StyleBuilder& border(BorderPosition pos, BorderStyle style, const Color& color = Color::BLACK);
StyleBuilder& allBorders(BorderStyle style, const Color& color = Color::BLACK);
StyleBuilder& topBorder(BorderStyle style, const Color& color = Color::BLACK);
StyleBuilder& bottomBorder(BorderStyle style, const Color& color = Color::BLACK);
StyleBuilder& leftBorder(BorderStyle style, const Color& color = Color::BLACK);
StyleBuilder& rightBorder(BorderStyle style, const Color& color = Color::BLACK);

enum class BorderStyle {
    None, Thin, Medium, Thick, Double, Dotted, 
    Dashed, DashDot, DashDotDot, SlantDashDot
};
```

### 填充设置

```cpp
// 填充模式
StyleBuilder& fill(PatternType pattern, const Color& fg_color, const Color& bg_color = Color::WHITE);
StyleBuilder& solidFill(const Color& color);
StyleBuilder& backgroundColor(const Color& color);  // 等同于 solidFill
StyleBuilder& pattern(PatternType pattern);

enum class PatternType {
    None, Solid, Gray50, Gray75, Gray25, HorizontalStripe,
    VerticalStripe, ReverseDiagonalStripe, DiagonalStripe,
    DiagonalCrosshatch, ThickDiagonalCrosshatch, 
    ThinHorizontalStripe, ThinVerticalStripe, 
    ThinReverseDiagonalStripe, ThinDiagonalStripe,
    ThinHorizontalCrosshatch, ThinDiagonalCrosshatch,
    Gray125, Gray0625
};
```

### 数字格式

```cpp
// 数字格式
StyleBuilder& numberFormat(const std::string& format_code);
StyleBuilder& numberFormat(int format_index);

// 常用格式快捷方法
StyleBuilder& generalFormat();
StyleBuilder& integerFormat();
StyleBuilder& decimalFormat(int decimals = 2);
StyleBuilder& percentFormat(int decimals = 2);
StyleBuilder& currencyFormat(const std::string& symbol = "$");
StyleBuilder& dateFormat(const std::string& format = "yyyy-mm-dd");
StyleBuilder& timeFormat(const std::string& format = "hh:mm:ss");
```

### 构建样式

```cpp
// 构建最终样式
FormatDescriptor build() const;

// 重置构建器
StyleBuilder& reset();

// 从现有格式创建
static StyleBuilder fromFormat(const FormatDescriptor& format);
```

---

## FormatDescriptor 格式描述符

FormatDescriptor 是不可变的格式对象，确保线程安全。

### 访问属性

```cpp
class FormatDescriptor {
public:
    // 属性访问
    const FontDescriptor& getFont() const;
    const AlignmentDescriptor& getAlignment() const;
    const BorderDescriptor& getBorder() const;
    const FillDescriptor& getFill() const;
    const NumberFormatDescriptor& getNumberFormat() const;
    const ProtectionDescriptor& getProtection() const;
    
    // 哈希和比较
    size_t hash() const;
    bool operator==(const FormatDescriptor& other) const;
    bool operator!=(const FormatDescriptor& other) const;
    
    // 调试信息
    std::string toString() const;
};
```

---

## Color 颜色系统

### 预定义颜色

```cpp
class Color {
public:
    // 常用颜色常量
    static const Color BLACK;
    static const Color WHITE;
    static const Color RED;
    static const Color GREEN;
    static const Color BLUE;
    static const Color YELLOW;
    static const Color CYAN;
    static const Color MAGENTA;
    static const Color DARK_BLUE;
    
    // 灰度色
    static const Color LIGHT_GRAY;
    static const Color DARK_GRAY;
    
    // 创建颜色
    static Color fromRGB(uint8_t r, uint8_t g, uint8_t b);
    static Color fromHex(const std::string& hex); // "#FF0000"
    static Color fromTheme(int theme_index, double tint = 0.0);
    static Color fromIndex(int color_index);
};
```

### 颜色操作

```cpp
// 颜色属性
uint8_t getRed() const;
uint8_t getGreen() const;
uint8_t getBlue() const;
uint8_t getAlpha() const;

// 格式转换
std::string toHex() const;
uint32_t toARGB() const;

// 颜色运算
Color lighter(double factor = 1.5) const;
Color darker(double factor = 0.7) const;
Color withAlpha(uint8_t alpha) const;
```

---

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
    
    class CommonUtils {
    public:
        static std::string cellReference(int row, int col);
        static std::pair<int, int> parseAddress(const std::string& address);
    };
}
```

### 路径处理

```cpp
class Path {
public:
    Path(const std::string& path);
    Path(const char* path);
    
    std::string string() const;
    std::string filename() const;
    std::string extension() const;
    
    bool exists() const;
    bool isFile() const;
    bool isDirectory() const;
    
    FILE* openForWrite(bool truncate = true) const;
    FILE* openForRead() const;
};
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

// 样式去重统计
struct DeduplicationStats {
    size_t total_requests;
    size_t unique_formats;
    double deduplication_ratio;
};

// FormatRepository方法
class FormatRepository {
public:
    DeduplicationStats getDeduplicationStats() const;
    size_t getFormatCount() const;
    std::shared_ptr<const FormatDescriptor> getFormat(int id) const;
    int addFormat(const FormatDescriptor& format);
};
```

---

## 📝 最佳实践

### 1. 正确的初始化流程
```cpp
// ✅ 正确的创建方式
auto workbook = std::make_unique<Workbook>(Path("output.xlsx"));
if (!workbook->open()) {
    throw std::runtime_error("无法创建工作簿");
}

// ❌ 错误：忘记调用open()
auto workbook = std::make_unique<Workbook>(Path("output.xlsx"));
// 直接使用会导致未初始化错误
```

### 2. 样式复用
```cpp
// ✅ 样式复用，减少内存使用
auto commonStyle = workbook->createStyleBuilder()
    .fontName("Calibri")
    .fontSize(11)
    .build();
int styleIndex = workbook->addFormat(commonStyle);

// 在多处复用
for (int i = 0; i < 1000; ++i) {
    auto formatDescriptor = workbook->getFormatDescriptor(styleIndex);
    worksheet->getCell(i, 0).setFormat(formatDescriptor);
}
```

### 3. 类型安全访问
```cpp
// ✅ 类型检查后访问
const auto& cell = worksheet->getCell(0, 0);
if (cell.isString()) {
    std::string value = cell.getStringValue();
} else if (cell.isNumber()) {
    double value = cell.getNumberValue();
}

// ❌ 不检查类型直接访问可能导致异常
```

### 4. 工作簿模式选择
```cpp
// ✅ 大数据处理使用批量模式
WorkbookOptions options;
options.mode = WorkbookMode::BATCH;
options.constant_memory = true;
workbook->setOptions(options);

// ✅ 实时处理使用流式模式  
options.mode = WorkbookMode::STREAMING;
```

### 5. 错误处理
```cpp
// ✅ 完善的错误处理
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

---

这个API参考文档基于 FastExcel v2.0 的当前源码结构，提供了完整而准确的API信息。

*最后更新: 2025-08-24*