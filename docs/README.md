# FastExcel 项目文档

FastExcel 是一个功能完整的现代 C++17 Excel 文件处理库，采用双架构设计，支持新旧API共存，专为高性能、大规模数据处理和完整Excel格式支持而设计。

## 📚 文档目录

> 💡 **快速导航**: 查看 **[文档索引](INDEX.md)** 获取完整的文档导航和分类

### 核心文档
- **[架构设计文档](architecture-design.md)** - 完整的项目架构设计和优化方案
- **[性能优化指南](performance-optimization-guide.md)** - 性能优化最佳实践和实施方案
- **[批量与流式架构详解](streaming-vs-batch-architecture-explained.md)** - 批量和流式模式的详细实现机制

### 实现文档
- **[XML 生成统一规范](xml-generation-guide.md)** - XML 生成的统一规范和实施指引
- **[主题实现指南](theme-implementation-guide.md)** - Excel 主题功能的实现指南
- **[共享公式优化路线图](shared-formula-optimization-roadmap.md)** - 共享公式系统的优化策略

## 🚀 快速开始

### 基本使用（新架构推荐）
```cpp
#include "fastexcel/FastExcel.hpp"

using namespace fastexcel;
using namespace fastexcel::core;

int main() {
    // 初始化库
    fastexcel::initialize();
    
    // 创建工作簿
    auto workbook = Workbook::create(Path("example.xlsx"));
    
    // 添加工作表
    auto worksheet = workbook->addWorksheet("数据表");
    
    // 写入各种数据类型
    worksheet->writeString(0, 0, "产品名称");
    worksheet->writeNumber(0, 1, 123.45);
    worksheet->writeBoolean(0, 2, true);
    worksheet->writeFormula(0, 3, "B1*2");
    
    // 创建样式
    auto style = workbook->createStyleBuilder()
        .font().name("微软雅黑").size(12).bold(true).color(Color::BLUE)
        .fill().pattern(PatternType::Solid).fgColor(Color::LIGHT_GRAY)
        .border().all(BorderStyle::Thin, Color::BLACK)
        .build();
    
    int styleId = workbook->addStyle(style);
    
    // 保存文件
    workbook->save();
    
    // 清理资源
    fastexcel::cleanup();
    return 0;
}
```

### 高级功能示例

#### 共享公式优化
```cpp
// 创建共享公式（优化大量相似公式）
int sharedFormulaId = worksheet->createSharedFormula(
    0, 2,    // 起始位置
    999, 2,  // 结束位置
    "A1*B1"  // 基础公式
);

// 自动优化现有公式
int optimizedCount = worksheet->optimizeFormulas();
auto report = worksheet->analyzeFormulaOptimization();
```

#### OPC包高性能编辑
```cpp
#include "fastexcel/opc/PackageEditor.hpp"

// 增量编辑现有Excel文件
auto editor = PackageEditor::open(Path("existing.xlsx"));
auto worksheet = editor->getWorkbook()->getWorksheet("Sheet1");
worksheet->writeNumber(5, 3, 99999.99);
editor->save(); // 只更新修改部分
```

#### 批量数据处理
```cpp
// 批量写入大数据集
std::vector<std::vector<double>> bigData(10000, std::vector<double>(50));
worksheet->writeRange(0, 0, bigData);

// 设置高性能模式
workbook->setHighPerformanceMode(true);
workbook->setAutoModeThresholds(1000000, 100*1024*1024); // 100万单元格，100MB内存
```

## 🏗️ 项目架构

FastExcel 采用现代 C++17 设计，具有双架构并存的创新设计：

### 🎯 双架构系统
- **新架构 2.0**：现代C++设计，不可变值对象，线程安全
- **旧架构兼容**：向后兼容API，支持平滑迁移

### 📦 主要组件

#### 核心模块 (Core)
- **Workbook** - 工作簿管理，文档属性，VBA项目支持
- **Worksheet** - 工作表功能，1100+行完整API
- **Cell** - 优化的24字节单元格结构，支持7种数据类型
- **SharedFormula** - 智能公式优化，内存节省50-80%
- **FormatDescriptor** - 不可变格式描述符，线程安全
- **StyleBuilder** - 流畅样式构建器，链式调用API

#### XML处理系统 (XML)
- **UnifiedXMLGenerator** - 统一XML生成架构
- **WorksheetXMLGenerator** - 工作表XML生成器
- **StyleSerializer** - 样式序列化器，支持完整Excel格式

#### 存储引擎 (Archive) 
- **ZipArchive** - 高性能ZIP处理
- **PackageEditor** - OPC包增量编辑
- **CompressionEngine** - 多种压缩算法支持

#### 读取系统 (Reader)
- **XLSXReader** - 完整XLSX文件解析器
- **StylesParser** - 样式解析，支持特殊图案如gray125
- **WorksheetParser** - 工作表数据解析

#### OPC包编辑 (OPC)
- **PackageEditor** - 增量编辑器
- **ZipRepackWriter** - 保真写回
- **PartGraph** - 依赖关系管理

#### 主题系统 (Theme)
- **Theme** - 完整Excel主题支持
- **ThemeParser** - 主题解析器
- **ThemeUtils** - 主题工具函数

#### 变更跟踪 (Tracking)
- **ChangeTrackerService** - 变更跟踪服务
- **DirtyManager** - 脏数据管理

## 📊 性能特性

### 🚄 极致性能优化
- **内存效率**: 24字节/Cell vs 传统64字节+，提升62%
- **样式去重**: 自动合并相同样式，节省50-80%存储
- **智能缓存**: LRU缓存系统，10x性能提升
- **并行处理**: 多线程支持，3-5x速度提升

### 📈 智能模式选择
- **AUTO模式**: 根据数据规模自动选择最优策略
- **BATCH模式**: 内存中构建，最佳压缩比
- **STREAMING模式**: 常量内存使用，处理无限大文件

### 🎯 高级优化技术
- **共享公式**: 自动检测相似公式模式，大幅节省内存
- **内存池**: 减少内存分配开销
- **零缓存流式**: 流式模式真正的常量内存
- **压缩优化**: 支持zlib-ng、libdeflate高性能压缩

## 🔧 完整功能支持

### 📝 数据类型
- ✅ **数字**: 整数、浮点、科学计数法
- ✅ **字符串**: Unicode、富文本、内联优化
- ✅ **布尔值**: TRUE/FALSE
- ✅ **公式**: 所有Excel函数、共享公式
- ✅ **日期时间**: 完整Excel日期系统
- ✅ **错误值**: #DIV/0!, #N/A, #VALUE!等
- ✅ **超链接**: URL、邮箱、内部引用

### 🎨 样式系统
- ✅ **字体**: 名称、大小、粗体、斜体、颜色
- ✅ **填充**: 纯色、渐变、图案（包括gray125）
- ✅ **边框**: 四边框线、对角线、多种样式
- ✅ **对齐**: 水平、垂直、换行、旋转、缩进
- ✅ **数字格式**: 内置格式、自定义格式代码
- ✅ **主题支持**: 完整Excel主题系统

### 📊 工作表功能
- ✅ **基本操作**: 添加、删除、重命名、移动、复制
- ✅ **合并单元格**: 合并、拆分、范围管理
- ✅ **自动筛选**: 数据筛选、高级筛选
- ✅ **冻结窗格**: 行列冻结、分割窗格
- ✅ **打印设置**: 打印区域、重复行列、页面设置
- ✅ **工作表保护**: 密码保护、选择性保护

### 🏢 工作簿管理
- ✅ **文档属性**: 标题、作者、公司、自定义属性
- ✅ **定义名称**: 命名区域、公式常量
- ✅ **VBA项目**: 保留和传输VBA代码
- ✅ **工作簿保护**: 结构保护、窗口保护

## 🔧 编译和安装

### 系统要求
- **C++17** 或更高版本
- **CMake 3.15+**
- **支持的编译器**: GCC 7+, Clang 6+, MSVC 2017+

### 构建步骤
```bash
# 配置项目
cmake -B cmake-build-debug -S .

# 编译项目
cmake --build cmake-build-debug

# 并行编译加速
cmake --build cmake-build-debug -j 4

# 运行示例
cd cmake-build-debug/bin/examples
./sheet_copy_with_format_example
```

### 构建选项
```bash
# 启用高性能压缩
cmake -B build -S . -DFASTEXCEL_USE_LIBDEFLATE=ON

# 完整开发环境
cmake -B build -S . \
  -DFASTEXCEL_BUILD_EXAMPLES=ON \
  -DFASTEXCEL_BUILD_TESTS=ON \
  -DFASTEXCEL_BUILD_UNIT_TESTS=ON
```

## 📝 丰富示例

项目包含20+个示例，覆盖各种使用场景：

### 基础功能
- `01_basic_usage.cpp` - 基本读写操作  
- `02_basic_usage.cpp` - 数据类型示例
- `03_reader_example.cpp` - 文件读取示例

### 格式和样式
- `04_formatting_example.cpp` - 完整样式演示
- `08_sheet_copy_with_format_example.cpp` - 带格式复制

### 高级功能
- `09_high_performance_edit_example.cpp` - 高性能编辑
- `20_new_edit_architecture_example.cpp` - 新架构演示
- `21_package_editor_test.cpp` - OPC包编辑器

### 性能测试
- `test_shared_formula.cpp` - 共享公式测试
- `test_package_editor.cpp` - 包编辑器测试

## 🧪 测试框架

项目使用 **GoogleTest** 框架，提供全面的测试覆盖：

### 测试类型
- **单元测试**: `test/unit/` - 核心功能测试
- **集成测试**: `test/integration/` - 组件协作测试  
- **性能测试**: `test/performance/` - 基准性能测试

### 运行测试
```bash
# 运行所有测试
cd cmake-build-debug && ctest -V

# 运行性能测试
./test/performance/benchmark_shared_formula
./test/performance/benchmark_xml_generation
```

## 📈 版本历史

- **v3.6** (当前) - 完整功能实现，双架构设计，性能全面优化
  - ✅ 双架构并存设计完成
  - ✅ 共享公式优化系统
  - ✅ OPC包增量编辑
  - ✅ 完整样式和主题支持
  - ✅ 高性能内存管理
  
- **v2.1.0** - 架构重构，统一XML生成，智能模式选择
- **v2.0.0** - 统一接口设计，性能优化，测试规范化
- **v1.2.0** - 流式处理支持
- **v1.0.0** - 初始版本

## 🤝 贡献指南

欢迎贡献代码！请遵循以下步骤：

1. **Fork 项目** 并创建功能分支
2. **遵循代码规范** - 使用现代C++17风格
3. **添加测试用例** - 确保新功能有完整测试
4. **更新文档** - 保持文档同步更新
5. **提交 Pull Request** - 详细描述修改内容

### 开发环境设置
```bash
# 完整开发环境
cmake -B build -S . \
  -DFASTEXCEL_BUILD_EXAMPLES=ON \
  -DFASTEXCEL_BUILD_TESTS=ON \
  -DFASTEXCEL_BUILD_UNIT_TESTS=ON \
  -DFASTEXCEL_BUILD_PERFORMANCE_TESTS=ON \
  -DCMAKE_BUILD_TYPE=Debug

# 构建并运行测试
cmake --build build --parallel 4
cd build && ctest -V --parallel 4
```

## 📄 许可证

本项目采用 **MIT 许可证**，允许商业和开源使用。详见 [LICENSE](../LICENSE) 文件。

## 🔗 相关链接

### 核心文档
- [架构设计详解](architecture-design.md) - 了解FastExcel的完整架构设计
- [性能优化指南](performance-optimization-guide.md) - 获取最佳性能的使用建议
- [批量与流式模式详解](streaming-vs-batch-architecture-explained.md) - 深入理解两种处理模式

### 实现指南  
- [XML生成统一规范](xml-generation-guide.md) - XML生成的技术规范
- [主题实现指南](theme-implementation-guide.md) - Excel主题功能实现
- [共享公式优化路线图](shared-formula-optimization-roadmap.md) - 公式优化策略

## 📋 当前开发状态

FastExcel 已完成完整的企业级功能实现：

### ✅ 已完成功能
- **双架构设计** - 新旧API完美并存
- **完整Excel支持** - 所有数据类型、样式、功能
- **高性能优化** - 内存、速度、压缩全面优化
- **共享公式系统** - 智能公式优化，大幅节省内存
- **OPC包编辑** - 增量修改，保持原始格式
- **主题系统** - 完整的Excel主题支持
- **变更跟踪** - 精确跟踪修改部分

### 🚧 持续改进
- **文档完善** - API参考手册生成
- **示例扩充** - 更多实战场景演示
- **性能调优** - 针对特定场景的进一步优化

### 🔮 未来规划
- **Web Assembly支持** - 浏览器环境运行
- **Python绑定** - Python语言接口
- **图表支持** - Excel图表和绘图对象

---

**FastExcel** - 企业级Excel处理解决方案，让数据处理更快、更简单、更专业！

*最后更新: 2025-08-10*