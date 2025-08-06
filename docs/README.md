# FastExcel 项目文档

FastExcel 是一个高性能的 C++ Excel 文件处理库，专注于快速生成和处理 XLSX 文件。

## 📚 文档目录

### 核心文档
- **[API 参考文档](FastExcel_API_Reference.md)** - 完整的 API 参考和使用指南
- **[性能优化指南](Performance_Optimization_Guide.md)** - 性能优化最佳实践
- **[架构分析](FastExcel_Architecture_Analysis.md)** - 项目架构设计文档

### 设计文档
- **[统一接口提案](Workbook_Unified_Interface_Proposal.md)** - 工作簿统一接口设计
- **[模式选择指南](Workbook_Mode_Selection_Guide.md)** - 工作簿模式选择建议
- **[设计模式改进](Design_Pattern_Improvements.md)** - 设计模式应用和改进
- **[架构图表](FastExcel_Architecture_Diagrams.md)** - 系统架构图表

### 实现文档
- **[XLSXReader 实现](XLSXReader_Implementation.md)** - XLSX 读取器实现细节

## 🚀 快速开始

### 基本使用
```cpp
#include "fastexcel/FastExcel.hpp"

int main() {
    // 初始化库
    fastexcel::initialize();
    
    // 创建工作簿
    auto workbook = fastexcel::core::Workbook::create("example.xlsx");
    workbook->open();
    
    // 添加工作表
    auto worksheet = workbook->addWorksheet("数据表");
    
    // 写入数据
    worksheet->writeString(0, 0, "Hello");
    worksheet->writeNumber(0, 1, 123.45);
    
    // 保存文件
    workbook->save();
    
    // 清理资源
    fastexcel::cleanup();
    return 0;
}
```

### 高级功能
- **批量写入**: 支持高效的批量数据写入
- **流式处理**: 支持大文件的流式处理，节省内存
- **格式化**: 完整的单元格格式化支持
- **多工作表**: 支持多个工作表管理
- **时间处理**: 统一的时间工具类

## 🏗️ 项目架构

FastExcel 采用现代 C++ 设计，主要组件包括：

- **Core**: 核心功能模块（Workbook, Worksheet, Format, Cell）
- **Archive**: ZIP 归档处理模块
- **XML**: XML 流式写入器
- **Utils**: 工具类（TimeUtils, Logger 等）
- **Reader**: XLSX 文件读取器

## 📊 性能特性

- **高性能**: 相比 libxlsxwriter 提升 20-40% 性能
- **内存优化**: 智能内存管理和缓存系统
- **流式处理**: 支持大文件的常量内存处理
- **批量操作**: 优化的批量数据写入

## 🔧 编译和安装

### 依赖项
- C++17 或更高版本
- CMake 3.15+
- 支持的编译器：GCC 7+, Clang 6+, MSVC 2019+

### 构建步骤
```bash
mkdir build && cd build
cmake ..
cmake --build .
```

### 测试
```bash
# 运行单元测试
ctest

# 运行示例
./examples/basic_usage
./examples/formatting_example
```

## 📝 示例代码

项目包含多个示例，展示不同的使用场景：

- `examples/basic_usage.cpp` - 基本使用示例
- `examples/formatting_example.cpp` - 格式化示例
- `examples/high_performance_edit_example.cpp` - 高性能编辑示例
- `examples/optimization_demo.cpp` - 优化演示
- `examples/reader_example.cpp` - 读取示例

## 🧪 测试

项目使用 Google Test 框架进行测试：

- `test/unit/` - 单元测试
- `test/integration/` - 集成测试

测试覆盖了核心功能、性能测试、兼容性测试等方面。

## 📈 版本历史

- **v1.3.0** (当前) - 统一接口设计，性能优化，测试规范化
- **v1.2.0** - 添加流式处理支持
- **v1.1.0** - 格式化功能完善
- **v1.0.0** - 初始版本

## 🤝 贡献指南

欢迎贡献代码！请遵循以下步骤：

1. Fork 项目
2. 创建功能分支
3. 提交更改
4. 添加测试用例
5. 更新文档
6. 提交 Pull Request

## 📄 许可证

本项目采用 MIT 许可证，详见 LICENSE 文件。

## 🔗 相关链接

- [项目主页](https://github.com/fastexcel/FastExcel)
- [问题反馈](https://github.com/fastexcel/FastExcel/issues)
- [API 文档](FastExcel_API_Reference.md)

---

**FastExcel** - 让 Excel 文件处理更快、更简单、更现代！