# FastExcel 批量与流式架构详解

## 概述

本文档详细解释 FastExcel 中批量模式和流式模式的具体实现机制，以及 `generateFileWithCallback()` 方法如何根据不同的 Writer 类型选择不同的处理策略。

## 架构层次

```
ExcelStructureGenerator (统一调度器)
    ↓
generateFileWithCallback() (智能分发)
    ↓
IFileWriter 接口 (策略模式)
    ↓
BatchFileWriter / StreamingFileWriter (具体实现)
```

## 核心机制：动态类型检测

### 1. generateFileWithCallback() 的智能分发

```cpp
bool ExcelStructureGenerator::generateFileWithCallback(const std::string& path, 
    std::function<void(const std::function<void(const char*, size_t)>&)> generator) {
    
    // 🔍 关键：运行时类型检测
    if (auto streaming_writer = dynamic_cast<StreamingFileWriter*>(writer_.get())) {
        // 流式模式处理
        return handleStreamingMode(path, generator);
    } else {
        // 批量模式处理
        return handleBatchMode(path, generator);
    }
}
```

### 2. 流式模式的具体处理

```cpp
// 流式模式：真正的零缓存写入
if (auto streaming_writer = dynamic_cast<StreamingFileWriter*>(writer_.get())) {
    // 步骤1：打开流式文件
    if (!writer_->openStreamingFile(path)) {
        return false;
    }
    
    // 步骤2：直接流式写入，不经过内存缓存
    generator([this](const char* data, size_t size) {
        writer_->writeStreamingChunk(data, size);  // 直接写入ZIP流
    });
    
    // 步骤3：关闭流式文件
    return writer_->closeStreamingFile();
}
```

**StreamingFileWriter 的实际实现**：
```cpp
// StreamingFileWriter::writeStreamingChunk()
bool StreamingFileWriter::writeStreamingChunk(const char* data, size_t size) {
    if (!streaming_file_open_) {
        LOG_ERROR("No streaming file is open");
        return false;
    }
    
    // 直接调用 FileManager 的流式写入
    bool success = file_manager_->writeStreamingChunk(data, size);
    
    if (success) {
        stats_.total_bytes += size;  // 仅更新统计
    }
    
    return success;
}
```

**流式模式的数据流向**：
```
XML生成器 → writeStreamingChunk() → FileManager::writeStreamingChunk() → ZIP压缩流 → 磁盘文件
(无中间缓存，直接写入)
```

### 3. 批量模式的具体处理

```cpp
// 批量模式：先缓存，后批量写入
else {
    // 步骤1：收集所有XML数据到内存
    std::string content;
    generator([&content](const char* data, size_t size) {
        content.append(data, size);  // 缓存到字符串
    });
    
    // 步骤2：一次性写入
    return writer_->writeFile(path, content);
}
```

**BatchFileWriter 的实际实现**：
```cpp
// BatchFileWriter::writeFile()
bool BatchFileWriter::writeFile(const std::string& path, const std::string& content) {
    // 仅收集到内存，不立即写入
    files_.emplace_back(path, content);
    stats_.batch_files++;
    stats_.total_bytes += content.size();
    
    LOG_DEBUG("Collected file for batch write: {} ({} bytes)", path, content.size());
    return true;  // 总是返回 true，实际写入在 flush() 时进行
}

// BatchFileWriter::flush() - 在 finalize() 时调用
bool BatchFileWriter::flush() {
    // 一次性写入所有收集的文件
    bool success = file_manager_->writeFiles(std::move(files_));
    
    if (success) {
        stats_.files_written = files_.size();
    }
    
    files_.clear();  // 清空缓存
    return success;
}
```

**批量模式的数据流向**：
```
XML生成器 → 内存缓存(vector<pair<path,content>>) → flush() → FileManager::writeFiles() → ZIP压缩 → 磁盘文件
(先全部缓存，最后一次性写入)
```

## 具体调用示例

### 示例1：生成 docProps/app.xml

```cpp
// 在 ExcelStructureGenerator::generateBasicFiles() 中
if (!generateFileWithCallback("docProps/app.xml",
    [this](const std::function<void(const char*, size_t)>& callback) {
        workbook_->generateDocPropsAppXML(callback);
    })) {
    return false;
}
```

**执行流程分析**：

1. **调用 generateFileWithCallback()**
2. **检查 writer_ 类型**：
   - 如果是 `StreamingFileWriter*` → 流式处理
   - 如果是 `BatchFileWriter*` → 批量处理

3. **流式模式执行**：
   ```cpp
   writer_->openStreamingFile("docProps/app.xml");
   workbook_->generateDocPropsAppXML([this](const char* data, size_t size) {
       writer_->writeStreamingChunk(data, size);  // 直接写入ZIP流
   });
   writer_->closeStreamingFile();
   ```

4. **批量模式执行**：
   ```cpp
   std::string content;
   workbook_->generateDocPropsAppXML([&content](const char* data, size_t size) {
       content.append(data, size);  // 先缓存到字符串
   });
   writer_->writeFile("docProps/app.xml", content);  // 添加到批量列表
   ```

### 示例2：批量模式的流式接口处理

BatchFileWriter 也实现了流式接口，但实际上是"伪流式"：

```cpp
// BatchFileWriter 的流式接口实现
bool BatchFileWriter::openStreamingFile(const std::string& path) {
    current_path_ = path;
    current_content_.clear();  // 准备缓存区
    streaming_file_open_ = true;
    return true;
}

bool BatchFileWriter::writeStreamingChunk(const char* data, size_t size) {
    current_content_.append(data, size);  // 仍然是缓存到内存
    return true;
}

bool BatchFileWriter::closeStreamingFile() {
    // 将缓存的内容添加到批量文件列表
    files_.emplace_back(current_path_, current_content_);
    
    // 清理临时状态
    current_path_.clear();
    current_content_.clear();
    streaming_file_open_ = false;
    return true;
}
```

**关键差异**：
- **StreamingFileWriter**: `writeStreamingChunk()` → 直接写入ZIP流
- **BatchFileWriter**: `writeStreamingChunk()` → 缓存到 `current_content_` 字符串

## Writer 类型的选择机制

### 在 ExcelStructureGenerator 构造时确定

```cpp
// 在 Workbook::generateWithGenerator() 中
std::unique_ptr<IFileWriter> writer;
if (use_streaming_writer) {
    writer = std::make_unique<StreamingFileWriter>(file_manager_.get());
} else {
    writer = std::make_unique<BatchFileWriter>(file_manager_.get());
}
ExcelStructureGenerator generator(this, std::move(writer));
```

### 智能选择逻辑

```cpp
// 在 ExcelStructureGenerator::generate() 中
bool use_streaming = false;

switch (options_.mode) {
    case WorkbookMode::AUTO:
        // 根据数据量自动选择
        if (total_cells > options_.auto_mode_cell_threshold ||
            estimated_memory > options_.auto_mode_memory_threshold) {
            use_streaming = true;
        }
        break;
    case WorkbookMode::STREAMING:
        use_streaming = true;
        break;
    case WorkbookMode::BATCH:
        use_streaming = false;
        break;
}
```

## 性能对比

### 内存使用

| 模式 | 基础文件内存 | 工作表内存 | 总内存 |
|------|-------------|------------|--------|
| 批量模式 | 所有文件缓存 | 所有数据缓存 | 高 |
| 流式模式 | 零缓存 | 逐行处理 | 低 |

### 处理速度

| 模式 | 小文件 | 大文件 | 压缩效率 |
|------|--------|--------|----------|
| 批量模式 | 快 | 慢(内存不足) | 高 |
| 流式模式 | 稍慢 | 快 | 中等 |

## 调试和验证

### 如何确认当前使用的模式

```cpp
// 在日志中查看
LOG_INFO("Starting Excel structure generation using {}", writer_->getTypeName());

// 输出示例：
// "Starting Excel structure generation using BatchFileWriter"
// "Starting Excel structure generation using StreamingFileWriter"
```

### 性能统计

```cpp
auto stats = generator.getWriterStats();
LOG_INFO("Files: {} (batch: {}, streaming: {})", 
         stats.files_written, stats.batch_files, stats.streaming_files);
```

## 完整执行流程图

### 流式模式完整流程

```
ExcelStructureGenerator::generate()
    ↓
generateBasicFiles()
    ↓
generateFileWithCallback("docProps/app.xml", generator)
    ↓
dynamic_cast<StreamingFileWriter*>(writer_.get()) ✓
    ↓
StreamingFileWriter::openStreamingFile("docProps/app.xml")
    ↓
file_manager_->openStreamingFile("docProps/app.xml")  // 打开ZIP流
    ↓
generator([this](const char* data, size_t size) {
    writer_->writeStreamingChunk(data, size);
})
    ↓
workbook_->generateDocPropsAppXML(callback)
    ↓ (多次调用)
StreamingFileWriter::writeStreamingChunk(data, size)
    ↓
file_manager_->writeStreamingChunk(data, size)  // 直接写入ZIP
    ↓
StreamingFileWriter::closeStreamingFile()
    ↓
file_manager_->closeStreamingFile()  // 完成文件写入
```

### 批量模式完整流程

```
ExcelStructureGenerator::generate()
    ↓
generateBasicFiles()
    ↓
generateFileWithCallback("docProps/app.xml", generator)
    ↓
dynamic_cast<StreamingFileWriter*>(writer_.get()) ✗
    ↓
std::string content;
generator([&content](const char* data, size_t size) {
    content.append(data, size);
});
    ↓
workbook_->generateDocPropsAppXML(callback)
    ↓ (多次调用)
content.append(data, size)  // 缓存到字符串
    ↓
BatchFileWriter::writeFile("docProps/app.xml", content)
    ↓
files_.emplace_back("docProps/app.xml", content)  // 添加到批量列表
    ↓
... (处理所有文件)
    ↓
finalize()
    ↓
BatchFileWriter::flush()
    ↓
file_manager_->writeFiles(std::move(files_))  // 一次性写入所有文件
```

## 内存使用对比

### 生成 10MB Excel 文件的内存使用情况

| 阶段 | 流式模式 | 批量模式 |
|------|----------|----------|
| 基础文件生成 | ~50KB | ~500KB (缓存) |
| 工作表生成 | ~100KB | ~9.5MB (缓存) |
| 最终化 | ~50KB | ~10MB (写入时) |
| **峰值内存** | **~150KB** | **~10MB** |

### 性能特点对比

| 特性 | 流式模式 | 批量模式 |
|------|----------|----------|
| 内存使用 | 常量级 O(1) | 线性增长 O(n) |
| 写入延迟 | 实时写入 | 延迟到最后 |
| 压缩效率 | 中等 | 最优 |
| 错误恢复 | 困难 | 容易 |
| 适用场景 | 大文件 | 小到中等文件 |

## 总结

`generateFileWithCallback()` 方法的核心价值在于：

1. **统一接口** - 上层调用代码完全相同
2. **智能分发** - 运行时根据 Writer 类型选择策略
3. **性能优化** - 流式模式实现真正的零缓存
4. **代码复用** - 消除了重复的处理逻辑
5. **透明切换** - 用户无需关心底层实现差异

这种设计实现了策略模式的完美应用，让批量和流式模式在统一的接口下各自发挥最佳性能。通过 `dynamic_cast` 的运行时类型检测，系统能够智能地选择最适合的处理策略，既保证了代码的简洁性，又实现了性能的最优化。