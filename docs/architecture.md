# FastExcel 项目架构设计文档

## 1. 项目概述

FastExcel 是一个高性能的C++17 Excel文件处理库，专注于提供高效的Excel文件读写功能。项目采用分层架构设计，遵循SOLID原则，支持多种工作模式，并提供了完整的Excel文件格式支持。

### 1.1 核心特性

- **高性能**: 支持批量模式和流式模式，自动选择最优处理策略
- **全功能**: 完整支持Excel文件格式，包括样式、公式、图表等
- **易用性**: 提供类似libxlsxwriter的API，同时支持现代C++特性
- **扩展性**: 模块化设计，便于功能扩展和维护
- **线程安全**: 核心组件支持多线程并发访问

### 1.2 技术规格

- **编程语言**: C++17
- **构建系统**: CMake 3.15+
- **支持格式**: XLSX (Excel 2007+)
- **压缩算法**: Zlib/Deflate，支持多种压缩引擎
- **XML处理**: 自定义高性能XML流式读写器

## 2. 架构概览

FastExcel采用分层架构设计，共分为7个主要层次：

```
┌─────────────────────────────────────┐
│           应用层 (Application)       │  ← 用户API接口
├─────────────────────────────────────┤
│           核心层 (Core)             │  ← 业务逻辑核心
├─────────────────────────────────────┤
│          读取器层 (Reader)          │  ← Excel文件解析
├─────────────────────────────────────┤
│        XML处理层 (XML)              │  ← XML序列化/反序列化
├─────────────────────────────────────┤
│          主题层 (Theme)             │  ← 主题和样式管理
├─────────────────────────────────────┤
│         压缩层 (Archive)            │  ← ZIP文件处理
├─────────────────────────────────────┤
│         工具层 (Utils)              │  ← 通用工具和日志
└─────────────────────────────────────┘
```

## 3. 模块详细设计

### 3.1 核心层 (Core)

核心层是整个系统的业务逻辑中心，负责Excel文档的数据模型和业务操作。

#### 3.1.1 主要类结构

```cpp
namespace fastexcel::core {
    class Workbook;           // 工作簿：整个Excel文件的根对象
    class Worksheet;          // 工作表：单个表格数据管理
    class Cell;              // 单元格：最小数据单元
    class Format;            // 格式：样式信息
    class FormatRepository;  // 格式仓储：样式去重管理
    class ExcelStructureGenerator; // Excel结构生成器
}
```

#### 3.1.2 Workbook类 - 工作簿管理器

**设计目标**: 作为Excel文件的根对象，管理整个工作簿的生命周期和全局配置。

**核心职责**:
- 工作表生命周期管理
- 全局样式和格式管理
- 文档属性和元数据管理  
- 文件保存和加载协调
- 工作模式智能选择

**关键特性**:
```cpp
class Workbook {
    // 智能模式选择
    enum class WorkbookMode {
        AUTO,      // 自动选择最优模式
        BATCH,     // 批量模式（小文件/低内存使用）
        STREAMING  // 流式模式（大文件/高性能）
    };
    
    // 格式管理 - 新架构
    std::unique_ptr<FormatRepository> format_repo_;
    
    // 智能脏数据管理
    std::unique_ptr<DirtyManager> dirty_manager_;
    
    // 工作模式选择器
    WorkbookModeSelector mode_selector_;
};
```

**设计模式应用**:
- **工厂模式**: 通过静态方法`create()`和`open()`创建实例
- **组合模式**: 管理多个Worksheet对象
- **策略模式**: 根据文件大小和内存使用选择处理模式

#### 3.1.3 Worksheet类 - 工作表管理器

**设计目标**: 管理单个工作表的所有数据和配置，支持高性能的单元格操作。

**核心职责**:
- 单元格数据存储和管理
- 行列格式设置
- 合并单元格管理
- 打印设置和视图配置
- 高性能数据写入优化

**优化特性**:
```cpp
class Worksheet {
    // 优化模式下的行缓存
    struct WorksheetRow {
        int row_num;
        std::map<int, Cell> cells;
        bool data_changed = false;
    };
    
    // 使用范围跟踪
    CellRangeManager range_manager_;
    
    // 内存池优化
    std::unique_ptr<MemoryPool> memory_pool_;
};
```

#### 3.1.4 Cell类 - 单元格优化存储

**设计目标**: 提供内存高效的单元格数据存储，支持多种数据类型。

**内存优化策略**:
```cpp
class Cell {
    // 位域压缩标志
    struct {
        CellType type : 4;
        bool has_format : 1;
        bool has_hyperlink : 1;
        bool has_formula_result : 1;
        uint8_t reserved : 1;
    } flags_;
    
    // Union节省内存
    union CellValue {
        double number;
        int32_t string_id;
        bool boolean;
        char inline_string[16];  // 短字符串内联存储
    } value_;
    
    // 延迟分配的扩展数据
    struct ExtendedData* extended_;
};
```

**优化亮点**:
- **内联字符串**: 16字节以内的字符串直接存储，避免堆分配
- **延迟分配**: 只有复杂单元格才分配扩展数据结构
- **位域压缩**: 使用位域压缩标志位，节省内存

#### 3.1.5 FormatRepository类 - 样式去重仓储

**设计目标**: 使用Repository模式实现线程安全的格式去重存储。

**架构优势**:
```cpp
class FormatRepository {
    // 线程安全保护
    mutable std::shared_mutex mutex_;
    
    // 不可变格式存储
    std::vector<std::shared_ptr<const FormatDescriptor>> formats_;
    
    // 哈希快速查找
    std::unordered_map<size_t, int> hash_to_id_;
    
    // 性能统计
    std::atomic<size_t> cache_hits_;
};
```

**设计原则应用**:
- **Repository模式**: 抽象数据访问层
- **不可变对象**: FormatDescriptor设计为不可变
- **线程安全**: 使用读写锁支持并发访问

### 3.2 读取器层 (Reader)

负责解析现有的Excel文件，将二进制格式转换为内存中的数据结构。

#### 3.2.1 XLSXReader类 - Excel文件解析器

**设计目标**: 高性能解析Excel文件，支持增量加载和错误恢复。

**核心架构**:
```cpp
class XLSXReader {
    // 系统层API：高性能，使用错误码
    core::ErrorCode open();
    core::ErrorCode loadWorkbook(std::unique_ptr<core::Workbook>& workbook);
    core::ErrorCode loadWorksheet(const std::string& name, 
                                 std::shared_ptr<core::Worksheet>& worksheet);
    
    // 解析器模块
    std::unique_ptr<StylesParser> styles_parser_;
    std::unique_ptr<SharedStringsParser> shared_strings_parser_;
    std::unique_ptr<WorksheetParser> worksheet_parser_;
};
```

**解析器职责分离**:
- **StylesParser**: 样式信息解析
- **SharedStringsParser**: 共享字符串表解析  
- **WorksheetParser**: 工作表数据解析
- **ContentTypesParser**: 内容类型解析

### 3.3 XML处理层 (XML)

提供高性能的XML序列化和反序列化功能，是整个系统性能的关键层。

#### 3.3.1 XMLStreamWriter类 - 流式XML写入器

**设计目标**: 参考libxlsxwriter实现极致性能的XML写入。

**性能优化**:
```cpp
class XMLStreamWriter {
    // 固定大小缓冲区
    static constexpr size_t BUFFER_SIZE = 8192;
    char buffer_[BUFFER_SIZE];
    
    // 高效转义算法
    void escapeAttributesToBuffer(const char* text, size_t length);
    void escapeDataToBuffer(const char* text, size_t length);
    
    // 直接文件写入模式
    void setDirectFileMode(FILE* file, bool take_ownership);
};
```

**优化策略**:
- **固定缓冲区**: 避免动态内存分配
- **预定义转义**: 使用memcpy和预定义长度
- **批量属性**: 减少系统调用次数

#### 3.3.2 XML生成服务体系

**统一XML生成架构**:
```cpp
namespace xml {
    class WorkbookXMLGenerator;    // 工作簿XML生成
    class StyleSerializer;         // 样式XML生成  
    class SharedStrings;          // 共享字符串XML生成
    class ContentTypes;           // 内容类型XML生成
    class Relationships;          // 关系XML生成
}
```

### 3.4 压缩层 (Archive)

管理ZIP文件的读写操作，支持多种压缩算法和并行处理。

#### 3.4.1 压缩引擎体系

**策略模式应用**:
```cpp
class CompressionEngine {
public:
    virtual ~CompressionEngine() = default;
    virtual bool compress(const void* src, size_t src_len,
                         void* dst, size_t& dst_len) = 0;
};

class ZlibEngine : public CompressionEngine;
class LibDeflateEngine : public CompressionEngine;
```

**压缩引擎选择**:
- **ZlibEngine**: 标准zlib压缩，兼容性好
- **LibDeflateEngine**: 高性能deflate实现，速度优先

#### 3.4.2 ZipArchive类 - ZIP文件管理

**组合模式应用**:
```cpp
class ZipArchive {
    std::unique_ptr<ZipReader> reader_;
    std::unique_ptr<ZipWriter> writer_;
    
    enum class Mode { None, Read, Write, ReadWrite } mode_;
};
```

### 3.5 OPC层 (Open Packaging Convention)

实现Excel文件包结构的编辑和管理功能。

#### 3.5.1 PackageEditor类 - 包编辑器

**设计目标**: 支持Excel文件的增量编辑，保留未修改部分，提升性能。

**服务化架构**:
```cpp
class PackageEditor {
    // 核心服务组件（依赖注入）
    std::unique_ptr<IPackageManager> package_manager_;
    std::unique_ptr<xml::IXMLGenerator> xml_generator_;
    std::unique_ptr<tracking::IChangeTracker> change_tracker_;
    
    // 智能变更管理
    void detectChanges();
    ChangeStats getChangeStats();
};
```

**设计优势**:
- **增量更新**: 只重新生成修改的部分
- **保真编辑**: 保留原文件的未修改部分
- **智能检测**: 自动识别需要更新的组件

### 3.6 主题层 (Theme)

管理Excel文件的主题、颜色方案和字体配置。

#### 3.6.1 Theme类体系

```cpp
namespace theme {
    class Theme;                  // 主题根对象
    class ThemeColorScheme;       // 颜色方案
    class ThemeFontScheme;        // 字体方案
    class ThemeParser;            // 主题解析器
}
```

### 3.7 工具层 (Utils)

提供通用工具、日志系统和性能监控功能。

#### 3.7.1 Logger类 - 高性能日志系统

**设计特性**:
```cpp
class Logger {
    enum class Level { TRACE, DEBUG, INFO, WARN, ERROR, CRITICAL, OFF };
    
    // 模板化日志接口
    template<typename... Args>
    void info(const std::string& fmt_str, Args&&... args);
    
    // 异步日志写入
    void log_to_file_async(const std::string& message);
};
```

## 4. 设计模式应用

### 4.1 创建型模式

#### 4.1.1 工厂模式
```cpp
// Workbook工厂方法
std::unique_ptr<Workbook> Workbook::create(const Path& path);
std::unique_ptr<Workbook> Workbook::open(const Path& path);
```

#### 4.1.2 建造者模式
```cpp
// 样式构建器
StyleBuilder builder = workbook.createStyleBuilder();
int style_id = builder.setFontName("Arial")
                     .setFontSize(12)
                     .setBold(true)
                     .setBackgroundColor(Color::LIGHT_GRAY)
                     .build();
```

### 4.2 结构型模式

#### 4.2.1 组合模式
```cpp
// Workbook组合多个Worksheet
class Workbook {
    std::vector<std::shared_ptr<Worksheet>> worksheets_;
};
```

#### 4.2.2 策略模式
```cpp
// 文件写入策略
class ExcelStructureGenerator {
    std::unique_ptr<IFileWriter> writer_;  // BatchFileWriter或StreamingFileWriter
};
```

### 4.3 行为型模式

#### 4.3.1 观察者模式
```cpp
// 变更通知
class DirtyManager {
    void notifyWorksheetChanged(const std::string& worksheet_name);
    void notifyStyleChanged(int style_id);
};
```

#### 4.3.2 模板方法模式
```cpp
// XML生成模板
class BaseXMLGenerator {
protected:
    virtual void generateHeader() = 0;
    virtual void generateBody() = 0;
    virtual void generateFooter() = 0;
    
public:
    std::string generate() {
        generateHeader();
        generateBody();
        generateFooter();
    }
};
```

## 5. 性能优化策略

### 5.1 内存优化

#### 5.1.1 对象池模式
```cpp
class MemoryPool {
    template<typename T>
    T* allocate();
    
    template<typename T>
    void deallocate(T* ptr);
};
```

#### 5.1.2 内联存储
```cpp
// Cell类中的短字符串内联存储
union CellValue {
    double number;
    char inline_string[16];  // 避免小字符串的堆分配
};
```

### 5.2 I/O优化

#### 5.2.1 缓冲区管理
```cpp
class XMLStreamWriter {
    static constexpr size_t BUFFER_SIZE = 8192;
    char buffer_[BUFFER_SIZE];  // 固定大小缓冲区
};
```

#### 5.2.2 并行处理
```cpp
class MinizipParallelWriter {
    std::vector<std::thread> worker_threads_;
    
    void processFileInParallel(const FileEntry& file);
};
```

### 5.3 算法优化

#### 5.3.1 格式去重
```cpp
class FormatRepository {
    // 使用哈希表快速查找重复格式
    std::unordered_map<size_t, int> hash_to_id_;
};
```

#### 5.3.2 智能模式选择
```cpp
class WorkbookModeSelector {
    WorkbookMode selectOptimalMode(size_t cell_count, size_t memory_usage);
};
```

## 6. 线程安全设计

### 6.1 读写锁应用
```cpp
class FormatRepository {
    mutable std::shared_mutex mutex_;
    
    // 读操作使用共享锁
    std::shared_ptr<const FormatDescriptor> getFormat(int id) const {
        std::shared_lock lock(mutex_);
        // 读取操作
    }
    
    // 写操作使用独占锁
    int addFormat(const FormatDescriptor& format) {
        std::unique_lock lock(mutex_);
        // 写入操作
    }
};
```

### 6.2 原子操作
```cpp
class Logger {
    std::atomic<Level> current_level_{Level::INFO};
    std::atomic<size_t> current_file_size_{0};
};
```

## 7. 错误处理机制

### 7.1 错误码系统
```cpp
namespace core {
    enum class ErrorCode {
        Ok,
        FileNotFound,
        InvalidFormat,
        MemoryError,
        IoError
    };
}

// 使用错误码而非异常
ErrorCode result = reader.loadWorkbook(workbook);
if (result != ErrorCode::Ok) {
    // 错误处理
}
```

### 7.2 异常安全保证
```cpp
class Cell {
    // 强异常安全保证
    Cell& operator=(const Cell& other) {
        Cell temp(other);  // 可能抛出异常
        swap(temp);        // 不抛出异常
        return *this;
    }
};
```

## 8. 扩展性设计

### 8.1 插件架构
```cpp
// 未来支持插件扩展
class IExcelPlugin {
public:
    virtual void processWorkbook(Workbook& workbook) = 0;
    virtual std::string getName() const = 0;
};
```

### 8.2 格式支持扩展
```cpp
// 支持新的Excel功能
class IFeatureHandler {
public:
    virtual void handleFeature(const std::string& feature_data) = 0;
};
```

## 9. 测试架构

### 9.1 单元测试覆盖
- **Core模块**: 95%+ 代码覆盖率
- **Reader模块**: 90%+ 代码覆盖率  
- **XML模块**: 85%+ 代码覆盖率

### 9.2 性能测试
```cpp
// 性能基准测试
class PerformanceBenchmark {
    void benchmarkLargeFileWriting();
    void benchmarkMemoryUsage();
    void benchmarkCompressionSpeed();
};
```

## 10. 部署和使用

### 10.1 编译配置
```cmake
# CMake配置
add_library(fastexcel STATIC
    src/fastexcel/core/Workbook.cpp
    src/fastexcel/core/Worksheet.cpp
    # ...
)

target_compile_features(fastexcel PUBLIC cxx_std_17)
```

### 10.2 使用示例
```cpp
#include "fastexcel/FastExcel.hpp"

// 基本使用
auto workbook = fastexcel::core::Workbook::create("output.xlsx");
auto worksheet = workbook->addWorksheet("Sheet1");
worksheet->writeString(0, 0, "Hello");
worksheet->writeNumber(0, 1, 42.0);
workbook->save();
```

## 11. 未来发展方向

### 11.1 功能扩展
- **公式引擎**: 支持Excel公式计算
- **图表支持**: 完整的图表创建和编辑
- **数据透视表**: 数据分析功能
- **宏支持**: VBA宏的处理

### 11.2 性能优化
- **SIMD优化**: 使用向量指令加速数据处理
- **GPU加速**: 大数据处理的GPU支持
- **分布式处理**: 超大文件的分布式处理

### 11.3 生态系统
- **Python绑定**: 提供Python接口
- **Web服务**: HTTP API服务
- **数据库集成**: 直接与数据库系统集成

---

**文档版本**: v1.0  
**最后更新**: 2025-01-09  
**维护者**: wuxianggujun  
**项目仓库**: FastExcel C++ Library