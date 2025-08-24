# FastExcel 架构概览

本文档描述了 FastExcel 的整体架构设计，核心设计理念和主要组件。

## 🎯 设计理念

### 1. 现代C++17优先
- **RAII资源管理**: 智能指针自动管理内存，零内存泄漏
- **类型安全**: 模板化API提供编译期类型检查
- **移动语义**: 充分利用C++11/17移动语义优化性能
- **constexpr**: 编译期常量计算，提升运行时性能

### 2. 高性能设计
- **内存优化**: Cell结构仅32字节，相比传统库节省50%+
- **缓存友好**: 紧凑内存布局，提升CPU缓存命中率
- **并行处理**: 内置线程池，充分利用多核CPU
- **延迟分配**: 按需分配内存，避免不必要的开销

### 3. 线程安全
- **不可变对象**: FormatDescriptor等核心对象不可变，天然线程安全
- **读写分离**: 读多写少的场景使用读写锁优化
- **原子操作**: 统计计数等使用原子操作避免锁开销

## 🏗️ 整体架构

```
┌─────────────────────────────────────────────────────────────┐
│                    FastExcel Library                        │
├─────────────────────────────────────────────────────────────┤
│                   Public API Layer                         │
│  ┌─────────────┐ ┌─────────────┐ ┌─────────────────────────┐ │
│  │  Workbook   │ │ Worksheet   │ │    StyleBuilder         │ │
│  │     API     │ │    API      │ │       API               │ │
│  └─────────────┘ └─────────────┘ └─────────────────────────┘ │
├─────────────────────────────────────────────────────────────┤
│                   Core Components                          │
│  ┌─────────────┐ ┌─────────────┐ ┌─────────────────────────┐ │
│  │    Cell     │ │FormatRepo   │ │   SharedString          │ │
│  │  Structure  │ │ sitory      │ │      Table              │ │
│  └─────────────┘ └─────────────┘ └─────────────────────────┘ │
├─────────────────────────────────────────────────────────────┤
│                 Processing Layer                           │
│  ┌─────────────┐ ┌─────────────┐ ┌─────────────────────────┐ │
│  │XML Stream   │ │ Archive     │ │      Theme              │ │
│  │  Writer     │ │ Manager     │ │     System              │ │
│  └─────────────┘ └─────────────┘ └─────────────────────────┘ │
├─────────────────────────────────────────────────────────────┤
│                Infrastructure                              │
│  ┌─────────────┐ ┌─────────────┐ ┌─────────────────────────┐ │
│  │Thread Pool  │ │Memory Pool  │ │     Logger              │ │
│  │             │ │             │ │                         │ │
│  └─────────────┘ └─────────────┘ └─────────────────────────┘ │
└─────────────────────────────────────────────────────────────┘
```

## 📦 核心模块

### 1. Core Module (核心模块)
**职责**: Excel文件的基本操作和数据结构

**核心组件**:
- **`Cell`**: 优化的单元格结构，支持内联字符串存储
- **`Worksheet`**: 工作表管理，支持大规模数据处理
- **`Workbook`**: 工作簿管理，统一的文件操作接口
- **`FormatRepository`**: 线程安全的样式去重存储
- **`StyleBuilder`**: 流畅API的样式构建器

**设计特点**:
```cpp
// Cell 内存优化设计
class Cell {
private:
    struct {
        CellType type : 4;           // 位域压缩
        bool has_format : 1;
        bool has_hyperlink : 1;
        bool has_formula_result : 1;
        bool is_shared_formula : 1;
    } flags_;                        // 1字节
    
    union CellValue {
        double number;               // 8字节
        char inline_string[16];      // 短字符串内联
    } value_;                        // 16字节
    
    std::unique_ptr<ExtendedData> extended_; // 延迟分配
};
```

### 2. Archive Module (存档模块)
**职责**: ZIP文件的压缩和解压缩处理

**核心组件**:
- **`ZipArchive`**: 高级ZIP文件操作接口
- **`CompressionEngine`**: 可插拔的压缩引擎抽象
- **`ZlibEngine`**: Zlib压缩实现
- **`LibDeflateEngine`**: 高性能libdeflate实现
- **`FileManager`**: 文件系统操作抽象

**优化策略**:
```cpp
// 多级压缩策略
enum class CompressionLevel {
    STORE = 0,      // 无压缩，最快
    FAST = 1,       // 快速压缩
    BALANCED = 6,   // 平衡模式（默认）
    BEST = 9        // 最佳压缩比
};
```

### 3. XML Module (XML模块)
**职责**: Excel XML格式的生成和解析

**核心组件**:
- **`XMLStreamWriter`**: 高性能流式XML写入
- **`XMLStreamReader`**: 流式XML解析
- **`StyleSerializer`**: 样式的XML序列化
- **`WorksheetXMLGenerator`**: 工作表XML生成

**性能优化**:
```cpp
// 预编译转义表
class XMLStreamWriter {
private:
    static constexpr std::array<std::string_view, 256> ESCAPE_TABLE = {
        // 预计算的转义字符映射
    };
    
public:
    void writeEscapedText(std::string_view text) {
        // 使用查找表，避免运行时字符串操作
    }
};
```

### 4. OPC Module (OPC包模块)
**职责**: Office Open XML包格式的编辑和管理

**核心组件**:
- **`PackageEditor`**: OPC包编辑器，支持增量修改
- **`ZipRepackWriter`**: ZIP重新打包写入器
- **`PartGraph`**: 部件依赖关系管理

**增量编辑技术**:
```cpp
// 保真写回 - 只修改变更部分
class PackageEditor {
public:
    // 只重写修改的部分，保持其他部分不变
    bool updateWorksheetOnly(int sheet_id, const std::string& xml_content);
    
    // 保留未修改的样式和主题
    bool preserveUnmodifiedStyles();
};
```

### 5. Reader Module (读取模块)
**职责**: Excel文件的解析和读取

**核心组件**:
- **`XLSXReader`**: 主读取器，协调各个解析器
- **`WorksheetParser`**: 工作表数据解析
- **`StylesParser`**: 样式信息解析
- **`SharedStringsParser`**: 共享字符串解析

### 6. Utils Module (工具模块)
**职责**: 通用工具和基础设施

**核心组件**:
- **`Logger`**: 高性能日志系统
- **`ThreadPool`**: 线程池管理
- **`MemoryPool`**: 内存池优化
- **`PerformanceMonitor`**: 性能监控

## 🔄 数据流设计

### 1. 写入流程

```
用户数据输入
    ↓
Cell.setValue()          ← 类型安全的模板API
    ↓
FormatRepository         ← 样式去重
    ↓
SharedStringTable        ← 字符串去重
    ↓
XMLStreamWriter          ← 高性能XML生成
    ↓
CompressionEngine        ← 可配置压缩
    ↓
ZipArchive              ← 最终Excel文件
```

### 2. 读取流程

```
Excel文件
    ↓
ZipArchive              ← ZIP解压缩
    ↓
XMLStreamReader         ← XML解析
    ↓
StylesParser           ← 样式解析
SharedStringsParser    ← 字符串解析
WorksheetParser        ← 工作表解析
    ↓
FormatRepository       ← 样式重建
SharedStringTable      ← 字符串表重建
    ↓
Worksheet & Cell       ← 重建对象模型
    ↓
用户API访问
```

## ⚡ 性能优化策略

### 1. 内存优化

**Cell结构优化**:
- 基础Cell: 32字节 (vs 传统64+字节)
- 短字符串内联: 15字符以下零额外分配
- 延迟分配: ExtendedData只在需要时创建

**样式去重**:
```cpp
// FormatRepository 自动去重
class FormatRepository {
private:
    std::unordered_map<size_t, int> hash_to_id_;  // 哈希查找
    std::vector<std::shared_ptr<const FormatDescriptor>> formats_;
    
public:
    int addFormat(const FormatDescriptor& format) {
        size_t hash = format.hash();
        auto it = hash_to_id_.find(hash);
        if (it != hash_to_id_.end()) {
            return it->second;  // 复用现有样式
        }
        // 创建新样式...
    }
};
```

### 2. 并发优化

**线程安全设计**:
```cpp
// 不可变对象 - 天然线程安全
class FormatDescriptor {
    const FontDescriptor font_;      // 不可变
    const AlignmentDescriptor align_; // 不可变
    // ... 所有字段都是const
    
public:
    // 只提供读取接口，无修改方法
    const FontDescriptor& getFont() const { return font_; }
};

// 读写锁优化读多写少场景
class FormatRepository {
private:
    mutable std::shared_mutex mutex_;  // 读写锁
    
public:
    int addFormat(const FormatDescriptor& format) {
        std::unique_lock lock(mutex_);  // 写锁
        // ...
    }
    
    std::shared_ptr<const FormatDescriptor> getFormat(int id) const {
        std::shared_lock lock(mutex_);  // 读锁
        // ...
    }
};
```

### 3. 缓存优化

**热点数据缓存**:
```cpp
// 无锁热点缓存
class FormatRepository {
private:
    std::array<std::atomic<const FormatDescriptor*>, 64> hot_cache_;
    
public:
    std::shared_ptr<const FormatDescriptor> getFormat(int id) const {
        // 热点路径: 无锁访问
        if (id < 64) {
            if (auto* cached = hot_cache_[id].load(std::memory_order_acquire)) {
                return std::shared_ptr<const FormatDescriptor>(cached, [](const FormatDescriptor*) {});
            }
        }
        
        // 冷路径: 正常查找
        std::shared_lock lock(mutex_);
        // ...
    }
};
```

## 🔧 扩展机制

### 1. 插件架构

```cpp
// 可插拔的压缩引擎
class CompressionEngine {
public:
    virtual ~CompressionEngine() = default;
    virtual std::vector<uint8_t> compress(const std::vector<uint8_t>& data) = 0;
    virtual std::vector<uint8_t> decompress(const std::vector<uint8_t>& data) = 0;
};

// 具体实现
class ZlibEngine : public CompressionEngine { /* ... */ };
class LibDeflateEngine : public CompressionEngine { /* ... */ };
```

### 2. 配置驱动

```cpp
// 运行时配置
struct WorkbookOptions {
    bool constant_memory = false;         // 常量内存模式
    WorkbookMode mode = WorkbookMode::AUTO; // 自动模式选择
    size_t row_buffer_size = 5000;       // 行缓冲大小
    int compression_level = 6;           // 压缩级别
    size_t xml_buffer_size = 4 * 1024 * 1024; // XML缓冲
};
```

### 3. 监控和诊断

```cpp
// 完整的性能统计
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
```

## 📊 架构优势

### 1. 性能优势
- **内存效率**: 50%+ 内存节省
- **处理速度**: 3-5x 性能提升  
- **并发能力**: 线程安全设计支持高并发
- **缓存友好**: 紧凑内存布局提升缓存命中率

### 2. 可维护性
- **模块化**: 清晰的职责分离
- **类型安全**: 编译期错误检查
- **异常安全**: RAII保证资源正确释放
- **测试友好**: 依赖注入支持单元测试

### 3. 可扩展性
- **插件机制**: 可插拔的组件设计
- **配置驱动**: 运行时行为可配置
- **版本兼容**: 新旧API共存

### 4. 可靠性
- **内存安全**: 智能指针消除内存泄漏
- **线程安全**: 不可变对象和读写锁
- **错误处理**: 完善的异常处理机制
- **资源管理**: RAII确保资源正确释放

## 🎯 设计哲学

### 1. 性能优先但不牺牲安全
- 使用现代C++特性提升性能
- 编译期优化 > 运行时优化
- 零成本抽象原则

### 2. 渐进式优化
- 普通场景下开箱即用
- 高性能场景下可深度调优
- 向后兼容保证平滑迁移

### 3. 用户体验导向
- 直观的API设计
- 丰富的错误信息
- 完善的文档和示例

这个架构设计确保了 FastExcel 在保持高性能的同时，具备良好的可维护性和可扩展性，为用户提供了企业级的Excel处理解决方案。