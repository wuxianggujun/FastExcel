# ZIP架构重构设计

## 目标
将现有的`ZipArchive`类拆分为职责单一、高性能的`ZipReader`和`ZipWriter`类，并保留`ZipArchive`作为高层封装。

## 架构设计

```
                    ┌─────────────┐
                    │ ZipArchive  │ (高层封装，向后兼容)
                    └──────┬──────┘
                           │ 组合
                ┌──────────┴──────────┐
                ▼                     ▼
        ┌─────────────┐       ┌─────────────┐
        │  ZipReader  │       │  ZipWriter  │
        └─────────────┘       └─────────────┘
                │                     │
                └──────────┬──────────┘
                           ▼
                    ┌─────────────┐
                    │   minizip   │ (底层库)
                    └─────────────┘
```

## 类设计

### 1. ZipReader (专注读取)
```cpp
namespace fastexcel::archive {

class ZipReader {
public:
    // 核心功能
    - 打开/关闭ZIP文件
    - 列出所有条目
    - 检查条目存在性
    - 读取条目内容（支持流式）
    - 获取条目信息（大小、CRC、时间等）
    
    // 高性能特性
    - 条目信息缓存（避免重复扫描）
    - 流式读取大文件
    - 并行读取支持（线程安全）
    - 原始压缩数据获取（用于快速复制）
    
    // 特殊功能
    - 支持密码保护的ZIP
    - 支持ZIP64格式
    - 支持各种压缩方法
};
```

### 2. ZipWriter (专注写入)
```cpp
namespace fastexcel::archive {

class ZipWriter {
public:
    // 核心功能
    - 创建/打开ZIP文件
    - 添加文件/目录
    - 设置压缩级别
    - 写入文件内容（支持流式）
    
    // 高性能特性
    - 批量写入优化
    - 流式写入大文件
    - 并行压缩支持
    - 原始数据写入（避免重压缩）
    
    // 特殊功能
    - 支持加密
    - 支持ZIP64格式
    - 支持自定义压缩方法
    - 防重复写入
};
```

### 3. ZipArchive (高层封装)
```cpp
namespace fastexcel::archive {

class ZipArchive {
private:
    std::unique_ptr<ZipReader> reader_;
    std::unique_ptr<ZipWriter> writer_;
    
public:
    // 保持原有接口不变，向后兼容
    bool open(bool create = true);
    bool close();
    
    // 委托给相应的类
    ZipError addFile(...) { return writer_->addFile(...); }
    ZipError extractFile(...) { return reader_->extractFile(...); }
    
    // 高级功能
    - 编辑模式（repack）
    - 事务支持
    - 原子操作
};
```

## 实现细节

### 流式读取实现
```cpp
class ZipReader {
public:
    // 流式读取接口
    class EntryStream {
    public:
        size_t read(void* buffer, size_t size);
        bool seek(int64_t offset, int whence);
        int64_t tell() const;
        bool eof() const;
    };
    
    std::unique_ptr<EntryStream> openStream(const std::string& path);
    
    // 回调式读取
    bool streamEntry(const std::string& path, 
                    std::function<bool(const uint8_t*, size_t)> callback);
};
```

### 流式写入实现
```cpp
class ZipWriter {
public:
    // 流式写入接口
    class EntryStream {
    public:
        size_t write(const void* data, size_t size);
        bool flush();
        bool finish();
    };
    
    std::unique_ptr<EntryStream> createStream(const std::string& path);
    
    // 从其他流复制
    bool copyFromStream(const std::string& path, std::istream& input);
};
```

### 高效复制实现
```cpp
// 直接复制压缩数据，避免解压再压缩
bool efficientCopy(ZipReader& source, ZipWriter& dest, 
                  const std::string& entry) {
    // 获取原始压缩数据
    auto rawData = source.getRawCompressedData(entry);
    auto info = source.getEntryInfo(entry);
    
    // 直接写入压缩数据
    return dest.writeRawEntry(info, rawData);
}
```

## 性能优化策略

### 1. 读取优化
- **条目缓存**：首次打开时缓存所有条目信息
- **延迟加载**：只在需要时才读取实际数据
- **并行读取**：支持多线程同时读取不同条目
- **内存映射**：对于大文件使用内存映射

### 2. 写入优化
- **批量写入**：减少中央目录更新次数
- **异步压缩**：使用线程池进行压缩
- **智能缓冲**：根据文件大小自动调整缓冲区
- **预分配空间**：避免频繁的文件扩展

### 3. 内存优化
- **流式处理**：避免一次性加载大文件
- **缓冲池**：复用缓冲区，减少内存分配
- **压缩级别自适应**：根据文件类型选择合适的压缩级别

## 使用示例

### 只读场景
```cpp
ZipReader reader("archive.zip");
reader.open();

// 流式读取大文件
auto stream = reader.openStream("large_file.bin");
char buffer[8192];
while (size_t n = stream->read(buffer, sizeof(buffer))) {
    process(buffer, n);
}
```

### 只写场景
```cpp
ZipWriter writer("output.zip");
writer.open();

// 流式写入大文件
auto stream = writer.createStream("large_file.bin");
for (const auto& chunk : data_chunks) {
    stream->write(chunk.data(), chunk.size());
}
stream->finish();
```

### 编辑场景（高效重打包）
```cpp
ZipReader source("original.xlsx");
ZipWriter dest("modified.xlsx");

source.open();
dest.open();

// 复制未修改的文件（高效，不解压）
for (const auto& entry : source.listEntries()) {
    if (!isModified(entry)) {
        dest.copyRawFrom(source, entry);
    }
}

// 写入修改的文件
dest.addFile("xl/worksheets/sheet1.xml", newContent);

dest.close();
source.close();
```

## 迁移计划

### 第一阶段：实现新类
1. 实现`ZipReader`类（基于现有的读取代码）
2. 实现`ZipWriter`类（基于现有的写入代码）
3. 保持`ZipArchive`不变

### 第二阶段：重构ZipArchive
1. 修改`ZipArchive`使用新的`ZipReader`和`ZipWriter`
2. 保持所有公共接口不变
3. 添加废弃标记引导用户使用新类

### 第三阶段：优化性能
1. 实现高级流式接口
2. 添加并行处理支持
3. 优化内存使用

## 优势总结

1. **职责清晰**：每个类只负责一个方面
2. **性能更优**：可以针对性优化
3. **使用灵活**：用户可以根据需求选择合适的类
4. **向后兼容**：保留`ZipArchive`，现有代码无需修改
5. **扩展性好**：易于添加新功能
6. **测试简单**：每个类可以独立测试

## 注意事项

1. **线程安全**：确保多线程环境下的正确性
2. **错误处理**：提供清晰的错误信息
3. **资源管理**：使用RAII确保资源正确释放
4. **兼容性**：支持各种ZIP变体和扩展
