# FastExcel 项目文档

FastExcel 是一个高性能的 C++ Excel 文件处理库，专注于快速生成和处理 XLSX 文件。

## 📚 文档目录

> 💡 **快速导航**: 查看 **[文档索引](INDEX.md)** 获取完整的文档导航和分类

### 核心文档
- **[架构设计文档](architecture-design.md)** - 完整的项目架构设计和优化方案
- **[性能优化指南](performance-optimization-guide.md)** - 性能优化最佳实践和实施方案
- **[批量与流式架构详解](streaming-vs-batch-architecture-explained.md)** - 批量和流式模式的详细实现机制

### 实现文档
- **[XML 生成统一规范](xml-generation-guide.md)** - XML 生成的统一规范和实施指引
- **[主题实现指南](theme-implementation-guide.md)** - Excel 主题功能的实现指南

## 🚀 快速开始

### 基本使用
```cpp
#include "fastexcel/FastExcel.hpp"

int main() {
    // 初始化库
    fastexcel::initialize();
    
    // 创建工作簿
    auto workbook = fastexcel::core::Workbook::create(fastexcel::core::Path("example.xlsx"));
    
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
- **智能模式选择**: 自动根据数据量选择批量或流式模式
- **统一 XML 生成**: 采用策略模式的统一 XML 生成架构
- **内存池管理**: 高效的内存分配和回收机制
- **并行处理**: 支持多线程并行处理大数据集
- **完整样式支持**: 支持 Excel 的完整样式和主题系统

## 🏗️ 项目架构

FastExcel 采用现代 C++17 设计，主要组件包括：

- **Core**: 核心功能模块（Workbook, Worksheet, Cell, ExcelStructureGenerator）
- **Archive**: ZIP 归档处理模块（FileManager, ZipArchive）
- **XML**: 统一 XML 生成系统（IFileWriter, BatchFileWriter, StreamingFileWriter）
- **Utils**: 工具类（Logger, TimeUtils, MemoryPool）
- **Reader**: XLSX 文件读取器（XLSXReader）

## 📊 性能特性

- **智能写入策略**: 根据数据量自动选择最优的写入模式
- **零缓存流式处理**: 流式模式实现真正的常量内存使用
- **内存优化**: Cell 类优化至 24 字节，支持内存池管理
- **并行处理**: 支持工作表级别的并行生成和处理

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

- **v2.1.0** (当前) - 修复编译错误，完善架构设计，统一 XML 生成，智能模式选择
- **v2.0.0** - 统一接口设计，性能优化，测试规范化
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

- [架构设计详解](architecture-design.md) - 了解 FastExcel 的完整架构设计
- [性能优化指南](performance-optimization-guide.md) - 获取最佳性能的使用建议
- [批量与流式模式详解](streaming-vs-batch-architecture-explained.md) - 深入理解两种处理模式

## 📋 当前开发状态

FastExcel 已完成重大架构优化，主要改进包括：

- ✅ 统一 XML 生成架构已完成
- ✅ 批量/流式模式智能选择已实现
- ✅ 编译错误修复和架构完善已完成
- ✅ 策略模式实现和性能优化已完成
- 📋 后续计划：内存池优化、并行处理、完整测试覆盖

项目现在具有稳定的架构基础和完善的文档体系。

---

**FastExcel** - 让 Excel 文件处理更快、更简单、更现代！
