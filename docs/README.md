# FastExcel 文档中心

FastExcel 是一个现代 C++17 高性能 Excel 文件处理库，专为大规模数据处理和高并发场景设计。

## 📚 文档导航

### 🚀 快速开始
- [快速入门指南](guides/quick-start.md) - 5分钟上手FastExcel
- [安装配置指南](guides/installation.md) - 详细的安装和配置说明
- [基础示例](examples/) - 完整的代码示例和最佳实践

### 📖 API 参考
- [核心API参考](api/) - 完整的API文档
- [Workbook API](api/workbook.md) - 工作簿操作接口
- [Worksheet API](api/worksheet.md) - 工作表操作接口
- [Cell API](api/cell.md) - 单元格操作接口
- [样式系统 API](api/style.md) - 样式和格式化接口

### 🏗️ 架构设计
- [整体架构](architecture/overview.md) - 系统架构概览
- [核心组件设计](architecture/core-components.md) - 核心模块详解
- [性能优化策略](architecture/performance.md) - 内存和性能优化
- [线程安全设计](architecture/thread-safety.md) - 并发安全机制

### 📘 开发指南
- [开发环境搭建](guides/development-setup.md) - 开发环境配置
- [编码规范](guides/coding-standards.md) - 代码规范和最佳实践
- [测试指南](guides/testing.md) - 单元测试和集成测试
- [调试和性能分析](guides/debugging.md) - 问题诊断和性能调优

### 💡 使用案例
- [大数据处理](examples/large-data.md) - 处理百万级数据
- [批量文件处理](examples/batch-processing.md) - 批量Excel文件操作
- [实时数据导出](examples/real-time-export.md) - 实时数据生成Excel
- [多线程应用](examples/multi-threading.md) - 高并发场景使用

## 🌟 核心特性

### ⚡ 高性能优化
- **内存效率**: Cell结构仅32字节，相比传统库节省50%+内存
- **短字符串优化**: 15字符以下字符串内联存储，零额外分配
- **样式去重**: 自动检测重复样式，节省50-80%样式存储空间
- **并行处理**: 内置线程池支持，充分利用多核CPU性能

### 🛡️ 现代C++设计
- **RAII资源管理**: 智能指针自动管理内存，零内存泄漏
- **类型安全**: 模板化API提供编译期类型检查
- **异常安全**: 完整的异常安全保证
- **移动语义**: 充分利用C++11/17移动语义优化性能

### 📊 完整Excel支持
- **数据类型**: 支持所有Excel数据类型（数字、字符串、公式、日期等）
- **样式系统**: 完整的字体、填充、边框、对齐、数字格式支持
- **高级功能**: 合并单元格、自动筛选、冻结窗格、工作表保护
- **图片支持**: PNG、JPEG、GIF、BMP等格式的图片插入

### 🔧 企业级特性
- **CSV集成**: 智能CSV解析和导出，支持多种编码和分隔符
- **增量编辑**: OPC包编辑技术，支持修改现有Excel文件
- **主题系统**: 完整的Excel主题和颜色方案支持
- **性能监控**: 内置性能统计和内存使用监控

## 📊 性能指标

| 指标 | FastExcel | 传统库 | 提升幅度 |
|------|-----------|---------|----------|
| 内存效率 | 32字节/Cell | 64字节+/Cell | **50% ↓** |
| 样式存储 | 自动去重 | 重复存储 | **50-80% ↓** |
| 处理速度 | 并行+缓存 | 单线程 | **3-5x ↑** |
| 文件大小 | 优化压缩 | 标准压缩 | **15-30% ↓** |

## 🏷️ 版本信息

- **当前版本**: 2.0.0
- **API版本**: 稳定版
- **C++标准**: C++17
- **构建系统**: CMake 3.15+
- **支持平台**: Windows, Linux, macOS

## 🤝 贡献和支持

- **问题反馈**: [GitHub Issues](https://github.com/wuxianggujun/FastExcel/issues)
- **功能请求**: [GitHub Discussions](https://github.com/wuxianggujun/FastExcel/discussions)
- **贡献代码**: 查看 [贡献指南](guides/contributing.md)
- **技术支持**: 查看 [FAQ](guides/faq.md) 或创建issue

## 📝 更新日志

### v2.0.0 (当前版本)
- ✅ 完整的现代C++17重构
- ✅ 新增FormatDescriptor不可变样式系统
- ✅ 新增StyleBuilder流畅API
- ✅ 优化Cell内存布局，实现50%内存节省
- ✅ 新增短字符串内联存储优化
- ✅ 完善线程安全设计

查看完整的 [更新日志](CHANGELOG.md)

---

**FastExcel** - 让Excel处理变得简单高效 ⚡

Copyright © 2024 FastExcel Team. All rights reserved.
