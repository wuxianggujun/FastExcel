# FastExcel API 文档索引（v2.0）

本目录提供 FastExcel 各模块的 API 文档、使用指南和最佳实践。基于当前源码结构（v2.0），确保文档与代码实现完全一致。

## 📚 文档概览

### 🚀 快速入门
- **[Quick_Reference.md](Quick_Reference.md)** - FastExcel API 快速参考，10分钟上手
- **[core-api.md](core-api.md)** - 完整的核心 API 详细文档

### 🏗️ 模块概览

FastExcel 采用分层架构设计，主要模块包括：

#### 核心层（core）
工作簿/工作表/单元格、样式系统、格式化、共享字符串等核心功能。

**主要类**:
- `Workbook` - 工作簿管理，文件 I/O
- `Worksheet` - 工作表操作，单元格管理
- `Cell` - 单元格数据和格式
- `StyleBuilder` - 样式构建器，链式API
- `FormatDescriptor` - 不可变格式描述符
- `FormatRepository` - 线程安全的样式仓储

#### XML 处理层（xml）
统一 XML 生成器、流式读写器、序列化等。

**主要类**:
- `XMLStreamWriter` - 高性能 XML 写入
- `XMLStreamReader` - 流式 XML 解析
- `WorksheetXMLGenerator` - 工作表 XML 生成
- `UnifiedXMLGenerator` - 统一 XML 生成器

#### 归档层（archive）
ZIP 读写与压缩编解码。

**主要类**:
- `ZipArchive` - ZIP 文件操作
- `ZipReader`/`ZipWriter` - ZIP 读写器
- `CompressionEngine` - 压缩引擎抽象

#### OPC 包编辑（opc）
面向 XLSX 包的编辑器、包管理、部件图。

**主要类**:
- `PackageEditor` - OPC 包编辑器
- `ZipRepackWriter` - ZIP 重新打包
- `IChangeTracker` - 变更跟踪

#### 工具层（utils）
日志、路径、地址解析、时间/通用工具、XML 工具。

**主要类**:
- `Logger` - 统一日志系统
- `Path` - 跨平台路径处理
- `CommonUtils` - 地址解析和工具函数
- `XMLUtils` - XML 处理工具

## 🎯 使用流程

### 1. 基本使用流程
```
创建 Workbook → open() → 添加 Worksheet → 设置数据和格式 → save()
```

### 2. 样式设置流程
```
createStyleBuilder() → 链式设置 → build() → addFormat() → 应用到单元格
```

### 3. 高性能处理流程
```
设置 WorkbookOptions → 选择模式 → 批量操作 → 样式复用
```

## ⚡ 关键特性

### 线程安全设计
- `FormatDescriptor` 不可变对象，天然线程安全
- `FormatRepository` 读写锁优化
- 快照机制避免并发问题

### 内存优化
- Cell 结构仅 32 字节
- 样式自动去重，节省 50-80% 内存
- 短字符串内联存储

### 高性能处理
- 三种模式：AUTO/BATCH/STREAMING
- XML 流式生成，零内存拷贝
- SIMD 优化的 XML 转义

## 📖 API 设计原则

### 1. 类型安全
- 强类型 API，编译期错误检查
- 类型安全的单元格值访问

### 2. RAII 资源管理
- 智能指针自动管理内存
- 异常安全保证

### 3. 现代 C++ 特性
- C++17 标准
- 移动语义优化性能
- constexpr 编译期计算

## 📝 最佳实践

### 1. 正确的初始化
```cpp
auto workbook = std::make_unique<Workbook>(Path("file.xlsx"));
if (!workbook->open()) {
    throw std::runtime_error("无法创建工作簿");
}
```

### 2. 样式复用
```cpp
auto style = workbook->createStyleBuilder()
    .fontName("Calibri")
    .fontSize(11)
    .build();
int styleId = workbook->addFormat(style);
// 复用 styleId...
```

### 3. 类型检查
```cpp
const auto& cell = worksheet->getCell(0, 0);
if (cell.isString()) {
    std::string value = cell.getStringValue();
}
```

## 🔗 相关资源

- [架构概览](../architecture/overview.md) - 系统设计说明
- [性能优化指南](../performance-optimization-guide.md) - 性能调优
- [使用示例](../examples-tutorial.md) - 完整示例代码

---

**FastExcel API 文档** - 基于 v2.0 源码结构，确保准确性和完整性

*最后更新: 2025-08-24*