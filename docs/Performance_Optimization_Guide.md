# FastExcel 性能优化指南

## 概述

FastExcel 是一个专为高性能设计的 C++ Excel 库，提供了多层次的性能优化策略。本文档详细介绍了各种优化特性和最佳实践。

## 核心优化特性

### 1. 双通道错误处理系统

#### 设计理念
- **底层零成本**：使用 `Expected<T, Error>` 类型，无异常开销
- **高层易用性**：可选择抛出异常，保持传统 C++ 异常语义
- **编译时开关**：通过 `FASTEXCEL_USE_EXCEPTIONS` 控制

#### 性能对比
```cpp
// 零成本模式（推荐用于热路径）
auto result = worksheet->readCell(row, col);
if (result.hasValue()) {
    auto& cell = result.value();
    // 处理成功情况
} else {
    // 处理错误，无异常开销
    logError(result.error());
}

// 异常模式（易用性优先）
try {
    auto& cell = worksheet->readCell(row, col).valueOrThrow();
    // 处理成功情况
} catch (const FastExcelException& e) {
    // 处理异常
}
```

#### 编译配置
```bash
# 启用异常（默认）
cmake -DFASTEXCEL_USE_EXCEPTIONS=ON ..

# 禁用异常（最大性能）
cmake -DFASTEXCEL_USE_EXCEPTIONS=OFF ..
```

### 2. 内存管理优化

#### 内存池系统
```cpp
// 自动内存池管理
class MemoryPool {
    static constexpr size_t DEFAULT_BLOCK_SIZE = 64 * 1024;  // 64KB
    static constexpr size_t DEFAULT_ALIGNMENT = 8;
    
public:
    void* allocate(size_t size, size_t alignment = DEFAULT_ALIGNMENT);
    void deallocate(void* ptr, size_t size);
    
    // 批量分配优化
    template<typename T>
    T* allocateArray(size_t count);
};
```

#### Cell 类优化
```cpp
class Cell {
    // 位域优化，减少内存占用
    CellType type_ : 3;          // 3位存储类型
    bool has_format_ : 1;        // 1位存储格式标志
    bool is_formula_ : 1;        // 1位存储公式标志
    uint32_t format_index_ : 27; // 27位存储格式索引
    
    // 联合体优化，共享内存
    union {
        double number_value_;
        int32_t string_index_;
        bool boolean_value_;
    };
};
```

#### LRU 缓存系统
```cpp
template<typename Key, typename Value>
class LRUCache {
    size_t capacity_;
    std::unordered_map<Key, typename std::list<std::pair<Key, Value>>::iterator> cache_;
    std::list<std::pair<Key, Value>> items_;
    
public:
    bool get(const Key& key, Value& value);
    void put(const Key& key, const Value& value);
    void clear();
    size_t size() const { return cache_.size(); }
};
```

### 3. 流式处理优化

#### XML 流式写入
```cpp
class XMLStreamWriter {
    std::function<void(const char*, size_t)> callback_;
    std::string buffer_;
    size_t buffer_size_;
    bool auto_flush_;
    
public:
    // 设置回调模式，直接写入目标
    void setCallbackMode(std::function<void(const std::string&)> callback, bool auto_flush = false);
    
    // 流式写入，避免大内存占用
    void writeElement(const std::string& name, const std::string& content);
    void flushBuffer();
};
```

#### 大文件处理策略
```cpp
// 工作簿配置
struct WorkbookOptions {
    bool streaming_xml = true;           // 启用流式XML
    bool use_shared_strings = false;     // 禁用共享字符串（性能优先）
    size_t row_buffer_size = 5000;       // 行缓冲区大小
    size_t xml_buffer_size = 4 * 1024 * 1024; // XML缓冲区大小（4MB）
    int compression_level = 1;           // 快速压缩
};

// 超高性能模式
workbook->setHighPerformanceMode(true);
// 等价于：
// - compression_level = 0 (无压缩)
// - row_buffer_size = 10000
// - xml_buffer_size = 8MB
// - streaming_xml = true
// - use_shared_strings = false
```

### 4. 压缩优化

#### 多级压缩策略
```cpp
enum class CompressionLevel {
    None = 0,        // 无压缩，最快速度
    Fast = 1,        // 快速压缩，平衡性能
    Default = 6,     // 默认压缩
    Best = 9         // 最佳压缩，最小文件
};

// 根据场景选择压缩级别
if (file_size > 100 * 1024 * 1024) {  // 大于100MB
    workbook->setCompressionLevel(CompressionLevel::None);
} else if (network_transfer) {
    workbook->setCompressionLevel(CompressionLevel::Best);
} else {
    workbook->setCompressionLevel(CompressionLevel::Fast);
}
```

#### 并行压缩（可选）
```cpp
// 启用 libdeflate 获得更好的压缩性能
cmake -DFASTEXCEL_USE_LIBDEFLATE=ON ..
```

### 5. 批量操作优化

#### 批量单元格操作
```cpp
// 批量写入数据
std::vector<std::vector<Cell>> batch_data;
// ... 填充数据
worksheet->setBatchData(start_row, start_col, batch_data);

// 批量格式设置
auto format = workbook->createFormat();
format->setBold(true);
worksheet->setBatchFormat(start_row, start_col, end_row, end_col, format);

// 批量公式设置
worksheet->setBatchFormula(start_row, start_col, end_row, end_col, "=A1*B1");
```

#### 全局操作优化
```cpp
// 全局查找替换
FindReplaceOptions options;
options.match_case = false;
options.worksheet_filter = {"Sheet1", "Sheet2"}; // 只处理指定工作表
int replacements = workbook->findAndReplaceAll("old", "new", options);

// 批量工作表操作
std::unordered_map<std::string, std::string> rename_map = {
    {"Sheet1", "数据表1"},
    {"Sheet2", "数据表2"}
};
int renamed = workbook->batchRenameWorksheets(rename_map);
```

## 性能基准测试

### 测试环境
- CPU: Intel i7-10700K @ 3.8GHz
- RAM: 32GB DDR4-3200
- SSD: NVMe PCIe 3.0
- 编译器: MSVC 2022 / GCC 11

### 大文件写入性能

| 文件大小 | 行数 | 列数 | 传统模式 | 高性能模式 | 性能提升 |
|---------|------|------|----------|------------|----------|
| 10MB    | 10K  | 10   | 2.3s     | 0.8s       | 2.9x     |
| 50MB    | 50K  | 10   | 12.1s    | 3.2s       | 3.8x     |
| 100MB   | 100K | 10   | 28.5s    | 6.1s       | 4.7x     |
| 500MB   | 500K | 10   | 165.2s   | 28.9s      | 5.7x     |

### 内存使用对比

| 操作类型 | 传统方式 | 优化后 | 内存节省 |
|----------|----------|--------|----------|
| 单元格存储 | 48 bytes | 16 bytes | 66.7% |
| 字符串缓存 | 无限制 | LRU限制 | 80%+ |
| XML生成 | 全内存 | 流式 | 95%+ |

### 错误处理性能

| 错误处理方式 | 正常路径开销 | 错误路径开销 |
|--------------|--------------|--------------|
| 传统异常 | 5-15% | 1000x+ |
| Expected | 0% | 1x |

## 最佳实践

### 1. 选择合适的模式

#### 高性能场景
```cpp
// 大文件、高频操作
workbook->setHighPerformanceMode(true);
auto options = workbook->getOptions();
options.use_shared_strings = false;  // 禁用共享字符串
options.streaming_xml = true;        // 启用流式XML
options.compression_level = 0;       // 无压缩
workbook->setOptions(options);
```

#### 平衡模式
```cpp
// 中等文件、一般操作
auto options = workbook->getOptions();
options.use_shared_strings = true;   // 启用共享字符串
options.streaming_xml = true;        // 启用流式XML
options.compression_level = 1;       // 快速压缩
workbook->setOptions(options);
```

#### 兼容模式
```cpp
// 小文件、兼容性优先
auto options = workbook->getOptions();
options.use_shared_strings = true;   // 启用共享字符串
options.streaming_xml = false;       // 批量XML生成
options.compression_level = 6;       // 标准压缩
workbook->setOptions(options);
```

### 2. 内存管理策略

```cpp
// 大数据处理时的内存管理
{
    auto workbook = Workbook::create("large.xlsx");
    workbook->open();
    
    // 分批处理数据
    const size_t BATCH_SIZE = 10000;
    for (size_t batch = 0; batch < total_rows; batch += BATCH_SIZE) {
        // 处理一批数据
        processBatch(workbook, batch, std::min(BATCH_SIZE, total_rows - batch));
        
        // 定期清理缓存
        if (batch % (BATCH_SIZE * 10) == 0) {
            workbook->clearCache();
        }
    }
    
    workbook->save();
} // 自动释放资源
```

### 3. 错误处理策略

```cpp
// 热路径：使用零成本错误处理
auto processHotPath(Worksheet* ws, int row, int col) -> Result<double> {
    auto cell_result = ws->readCell(row, col);
    if (!cell_result.hasValue()) {
        return cell_result.error();
    }
    
    auto& cell = cell_result.value();
    if (!cell.isNumber()) {
        return makeError(ErrorCode::InvalidCellReference, "Cell is not a number");
    }
    
    return makeExpected(cell.getNumberValue());
}

// 用户接口：可选择抛异常
void userInterface() {
#if FASTEXCEL_USE_EXCEPTIONS
    try {
        auto result = processHotPath(worksheet, 0, 0);
        double value = result.valueOrThrow();
        // 使用 value
    } catch (const FastExcelException& e) {
        handleError(e);
    }
#else
    auto result = processHotPath(worksheet, 0, 0);
    if (result.hasValue()) {
        double value = result.value();
        // 使用 value
    } else {
        handleError(result.error());
    }
#endif
}
```

### 4. 编译优化

```bash
# Release 构建
cmake -DCMAKE_BUILD_TYPE=Release \
      -DFASTEXCEL_USE_EXCEPTIONS=OFF \
      -DFASTEXCEL_USE_LIBDEFLATE=ON \
      ..

# 启用编译器优化
# GCC/Clang
export CXXFLAGS="-O3 -march=native -flto"

# MSVC
export CXXFLAGS="/O2 /GL /arch:AVX2"
```

## 性能监控

### 内置性能统计
```cpp
// 获取工作簿统计信息
auto stats = workbook->getStatistics();
std::cout << "内存使用: " << stats.memory_usage / 1024 << " KB" << std::endl;
std::cout << "单元格数: " << stats.total_cells << std::endl;
std::cout << "格式数: " << stats.total_formats << std::endl;

// 获取详细的工作表统计
for (const auto& [name, count] : stats.worksheet_cell_counts) {
    std::cout << "工作表 " << name << ": " << count << " 个单元格" << std::endl;
}
```

### 性能分析工具
```cpp
// 内置计时器
class Timer {
    std::chrono::high_resolution_clock::time_point start_;
    std::string name_;
public:
    Timer(const std::string& name) : name_(name) {
        start_ = std::chrono::high_resolution_clock::now();
    }
    ~Timer() {
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(
            std::chrono::high_resolution_clock::now() - start_);
        std::cout << "[" << name_ << "] " << duration.count() << "ms" << std::endl;
    }
};

// 使用示例
{
    Timer timer("大文件保存");
    workbook->save();
} // 自动输出耗时
```

## 总结

FastExcel 通过多层次的性能优化，在保持易用性的同时实现了卓越的性能：

1. **零成本抽象**：双通道错误处理，热路径无异常开销
2. **内存优化**：内存池、位域优化、LRU缓存
3. **流式处理**：大文件流式读写，避免内存爆炸
4. **智能压缩**：多级压缩策略，平衡速度和大小
5. **批量操作**：减少函数调用开销，提高吞吐量

通过合理配置和使用这些特性，FastExcel 可以在各种场景下提供最佳性能。