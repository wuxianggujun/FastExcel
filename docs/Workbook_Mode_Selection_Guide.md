
# FastExcel 工作簿模式选择指南

## 概述

FastExcel 提供了灵活的模式选择机制，允许用户根据具体需求选择最合适的文件生成方式。系统支持三种模式：

1. **AUTO（自动模式）** - 根据数据量自动选择最佳模式
2. **BATCH（批量模式）** - 强制使用批量写入，适合小文件
3. **STREAMING（流式模式）** - 强制使用流式写入，适合大文件

## 问题背景

之前的实现存在以下问题：
- `streaming_xml` 默认为 `true`，导致总是使用流式模式
- 批量模式难以触发，需要手动设置 `streaming_xml = false`
- 自动选择逻辑只在两个标志都为 `false` 时才生效

## 新的解决方案

### 1. WorkbookMode 枚举

```cpp
enum class WorkbookMode {
    AUTO = 0,      // 自动选择（默认）
    BATCH = 1,     // 强制批量模式
    STREAMING = 2  // 强制流式模式
};
```

### 2. 使用方法

#### 基本使用
```cpp
auto workbook = Workbook::create("example.xlsx");
workbook->open();

// 设置模式
workbook->setMode(WorkbookMode::AUTO);      // 自动选择（默认）
workbook->setMode(WorkbookMode::BATCH);     // 强制批量模式
workbook->setMode(WorkbookMode::STREAMING); // 强制流式模式
```

#### 自定义自动模式阈值
```cpp
// 设置自动模式的阈值
workbook->setAutoModeThresholds(
    500000,              // 单元格数量阈值（50万）
    50 * 1024 * 1024    // 内存阈值（50MB）
);
```

### 3. 模式选择逻辑

#### AUTO 模式（默认）
- 如果单元格数量 > 阈值（默认100万）或预估内存 > 阈值（默认100MB），使用流式模式
- 否则使用批量模式

#### BATCH 模式
- 强制使用批量写入
- 所有文件内容先在内存中生成，然后一次性写入ZIP
- 优点：速度快，适合小文件
- 缺点：内存占用大

#### STREAMING 模式
- 强制使用流式写入
- 文件内容边生成边写入ZIP
- 优点：内存占用小，适合大文件
- 缺点：速度相对较慢

### 4. 向后兼容性

为了保持向后兼容，旧的 `setStreamingXML()` 方法仍然可用：

```cpp
// 旧API（已废弃，但仍可用）
workbook->setStreamingXML(false); // 相当于 setMode(WorkbookMode::BATCH)
workbook->setStreamingXML(true);  // 相当于 setMode(WorkbookMode::STREAMING)
```

### 5. 最佳实践

#### 小文件（< 100万单元格）
```cpp
// 使用默认的AUTO模式即可，会自动选择批量模式
auto workbook = Workbook::create("small_file.xlsx");
workbook->open();
// 不需要设置模式，AUTO会自动选择
```

#### 大文件（> 100万单元格）
```cpp
// AUTO模式会自动选择流式模式，或手动强制
auto workbook = Workbook::create("large_file.xlsx");
workbook->open();
workbook->setMode(WorkbookMode::STREAMING); // 可选，AUTO也会选择流式
```

#### 需要精确控制
```cpp
auto workbook = Workbook::create("custom_file.xlsx");
workbook->open();

// 根据具体情况选择
if (need_fast_generation && have_enough_memory) {
    workbook->setMode(WorkbookMode::BATCH);
} else if (need_low_memory) {
    workbook->setMode(WorkbookMode::STREAMING);
} else {
    // 让系统自动决定
    workbook->setMode(WorkbookMode::AUTO);
}
```

### 6. 性能对比

| 模式 | 速度 | 内存占用 | 适用场景 |
|------|------|----------|----------|
| BATCH | 快 | 高 | 小文件、内存充足 |
| STREAMING | 慢 | 低 | 大文件、内存受限 |
| AUTO | 自适应 | 自适应 | 通用场景 |

### 7. 注意事项

1. **constant_memory 选项**：如果设置了 `constant_memory = true`，会强制使用流式模式，无论 mode 设置为什么
2. **ZIP压缩**：压缩级别设置（`compression_level`）对两种模式都有效
3. **共享字符串**：`use_shared_strings` 选项对两种模式都有效

### 8. 示例代码

完整示例见 `examples/test_mode_selection.cpp`

## 总结

新的模式选择机制提供了更灵活的控制方式：
- 默认的 AUTO 模式能够智能选择最佳方案
- 可以根据需要强制使用特定模式
- 保持了向后兼容性
- 提供了自定义阈值的能力

这解决了之前"流式模式总是启动"的问题，让用户能够更好地控制文件生成行为。
