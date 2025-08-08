# FastExcel - 高性能 C++ Excel 文件处理库

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![C++17](https://img.shields.io/badge/C%2B%2B-17-blue.svg)](https://en.wikipedia.org/wiki/C%2B%2B17)
[![Build Status](https://img.shields.io/badge/build-passing-brightgreen.svg)]()

FastExcel 是一个现代化的 C++ Excel 文件处理库，专为高性能、内存效率和易用性而设计。支持完整的 XLSX 文件读取、写入和编辑功能。

## 🚀 核心特性

- **🔥 高性能**: 优化的内存管理和缓存系统，处理大文件时性能卓越
- **📝 完整编辑**: 支持单元格、行、列和工作表级别的全面编辑操作
- **🎨 样式支持**: 完整的 Excel 样式解析和应用（字体、颜色、边框、对齐等）
- **🛡️ 异常安全**: 完善的错误处理和异常管理系统
- **💾 内存优化**: 内存池和 LRU 缓存系统，最小化内存占用
- **🔒 线程安全**: 关键操作的线程安全保证
- **📊 批量操作**: 高效的批量数据处理和操作
- **🔍 搜索替换**: 强大的全局搜索和替换功能

## 📦 快速开始

### 基本用法

```cpp
#include "fastexcel/core/Workbook.hpp"

using namespace fastexcel;

int main() {
    try {
        // 🚀 直接打开XLSX文件进行编辑（如果文件不存在会自动创建）
        auto workbook = core::Workbook::open("data.xlsx");
        if (!workbook) {
            workbook = core::Workbook::create("data.xlsx");
            workbook->open();
        }
        
        // 获取或创建工作表
        auto sheet = workbook->getWorksheet("Sheet1");
        if (!sheet) {
            sheet = workbook->addWorksheet("Sheet1");
        }
        
        // 直接编辑数据
        sheet->writeString(0, 0, "姓名");
        sheet->writeString(0, 1, "年龄");
        sheet->writeString(1, 0, "张三");
        sheet->writeNumber(1, 1, 25);
        
        // 编辑现有单元格
        sheet->editCellValue(1, 1, 26.0); // 修改年龄
        
        // 查找和替换
        sheet->findAndReplace("张三", "李四", false, false);
        
        // 自动保存
        workbook->save();
        
        std::cout << "文件编辑完成！" << std::endl;
        
    } catch (const core::FastExcelException& e) {
        std::cerr << "错误: " << e.getDetailedMessage() << std::endl;
    }
    
    return 0;
}
```

## 🏗️ 项目结构

```
FastExcel/
├── src/fastexcel/
│   ├── core/                 # 核心功能
│   │   ├── Workbook.hpp/cpp  # 工作簿管理
│   │   ├── Worksheet.hpp/cpp # 工作表操作
│   │   ├── Cell.hpp/cpp      # 单元格数据
│   │   ├── Format.hpp/cpp    # 格式化支持
│   │   ├── MemoryPool.hpp/cpp# 内存池管理
│   │   ├── CacheSystem.hpp/cpp# 缓存系统
│   │   └── Exception.hpp/cpp # 异常处理
│   ├── reader/               # 文件读取
│   │   ├── XLSXReader.hpp/cpp# XLSX读取器
│   │   ├── WorksheetParser.hpp/cpp# 工作表解析
│   │   ├── StylesParser.hpp/cpp# 样式解析
│   │   └── SharedStringsParser.hpp/cpp# 共享字符串解析
│   ├── xml/                  # XML处理
│   │   └── StreamingXMLWriter.hpp/cpp# 流式XML写入
│   ├── archive/              # 压缩文件处理
│   │   └── ZipArchive.hpp/cpp# ZIP文件管理
│   └── compat/               # 兼容性层
│       └── libxlsxwriter.hpp # libxlsxwriter兼容
├── examples/                 # 示例代码
│   ├── basic_usage.cpp       # 基本使用示例
│   ├── formatting_example.cpp# 格式化示例
│   ├── comprehensive_formatting_test.cpp # 综合格式化测试
│   └── large_data_example.cpp# 大数据处理示例
├── test/                     # 测试用例
├── docs/                     # 文档
│   ├── FastExcel_Optimization_and_Testing_Summary.md # 优化总结
│   ├── Design_Pattern_Improvements.md # 设计模式建议
│   └── Code_Optimization_Summary.md # 代码优化总结
└── CMakeLists.txt           # 构建配置
```

## 🔧 编译和安装

### 系统要求

- C++17 或更高版本
- CMake 3.15+
- 支持的编译器：GCC 7+, Clang 6+, MSVC 2019+

### 使用 CMake 构建

```bash
git clone https://github.com/your-repo/FastExcel.git
cd FastExcel
mkdir build && cd build
cmake ..
make -j$(nproc)
```

### 集成到项目

#### CMake

```cmake
find_package(FastExcel REQUIRED)
target_link_libraries(your_target FastExcel::FastExcel)
```

#### 手动编译

```bash
g++ -std=c++17 -I/path/to/fastexcel/include \
    your_code.cpp -L/path/to/fastexcel/lib -lfastexcel
```

## 📚 功能特性详解

### 🎨 全面的格式化支持

FastExcel 提供了完整的 Excel 格式化功能：

```cpp
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/Format.hpp"

// 创建工作簿和工作表
core::Workbook workbook("formatted_example.xlsx");
auto worksheet = workbook.addWorksheet("格式化示例");

// 创建各种格式
auto titleFormat = workbook.addFormat();
titleFormat->setFontName("Arial");
titleFormat->setFontSize(16);
titleFormat->setBold(true);
titleFormat->setFontColor(0x0000FF);  // 蓝色
titleFormat->setBackgroundColor(0xD3D3D3);  // 浅灰色
titleFormat->setAlign(core::Format::Align::CENTER);
titleFormat->setBorder(core::Format::Border::THIN);

// 应用格式
worksheet->writeString(0, 0, "标题", titleFormat);

// 数字格式
auto currencyFormat = workbook.addFormat();
currencyFormat->setNumberFormat("¥#,##0.00");
worksheet->writeNumber(1, 0, 12345.67, currencyFormat);

// 日期格式
auto dateFormat = workbook.addFormat();
dateFormat->setNumberFormat("yyyy-mm-dd");
worksheet->writeNumber(2, 0, 45000, dateFormat);  // Excel日期序列号

workbook.close();
```

支持的格式化功能：
- **字体样式**：字体名称、大小、粗体、斜体、下划线、颜色
- **对齐方式**：水平对齐（左、中、右、填充）、垂直对齐（上、中、下）
- **边框样式**：无边框、细边框、中等边框、粗边框、虚线、点线
- **背景颜色**：支持RGB颜色值
- **数字格式**：货币、百分比、日期、自定义格式
- **文本处理**：文本换行、缩进设置
- **单元格操作**：合并单元格、行高、列宽设置

### 1. 高性能读写

```cpp
// 启用高性能模式
workbook->setHighPerformanceMode(true);

// 批量写入数据
std::vector<std::vector<std::string>> data = {
    {"张三", "25", "技术部"},
    {"李四", "30", "销售部"},
    // ... 更多数据
};

for (size_t row = 0; row < data.size(); ++row) {
    for (size_t col = 0; col < data[row].size(); ++col) {
        sheet->writeString(row, col, data[row][col]);
    }
}
```

### 2. 完整的编辑功能

```cpp
// 单元格编辑
sheet->editCellValue(1, 1, "新值", true); // 保留格式

// 复制和移动
sheet->copyCell(0, 0, 1, 1, true);        // 复制单元格
sheet->moveRange(0, 0, 2, 2, 5, 5);       // 移动范围

// 插入和删除
sheet->insertRows(1, 3);                  // 插入3行
sheet->deleteColumns(2, 1);               // 删除1列

// 搜索和替换
int count = sheet->findAndReplace("旧值", "新值", false, false);
```

### 3. 样式和格式

```cpp
// 创建格式
auto format = std::make_shared<core::Format>();
format->setFontName("Arial");
format->setFontSize(12);
format->setBold(true);
format->setBackgroundColor({255, 255, 0}); // 黄色背景

// 应用格式
sheet->writeString(0, 0, "标题", format);
```

### 4. 批量操作

```cpp
// 批量重命名工作表
std::unordered_map<std::string, std::string> rename_map = {
    {"Sheet1", "数据表"},
    {"Sheet2", "统计表"}
};
workbook->batchRenameWorksheets(rename_map);

// 全局搜索替换
core::Workbook::FindReplaceOptions options;
options.case_sensitive = false;
int replacements = workbook->findAndReplaceAll("旧公司", "新公司", options);
```

### 5. 内存和性能优化

```cpp
// 获取内存统计
auto& memory_manager = core::MemoryManager::getInstance();
auto stats = memory_manager.getGlobalStatistics();
std::cout << "内存使用: " << stats.total_memory_in_use << " bytes" << std::endl;

// 缓存统计
auto& cache_manager = core::CacheManager::getInstance();
auto cache_stats = cache_manager.getGlobalStatistics();
std::cout << "缓存命中率: " << cache_stats.string_cache.hit_rate() << std::endl;
```

## 🧪 测试和示例

### 运行测试套件

```bash
cd build
ctest --verbose
```

或者运行特定测试：

```bash
./test/unit/test_main
./test/integration/test_integration
```

### 综合格式化功能测试

我们提供了一个全面的格式化功能测试示例：

```bash
# 编译并运行综合格式化测试
cd build
make comprehensive_formatting_test
./examples/comprehensive_formatting_test
```

该测试程序包含：
- **基本格式化**：字体、颜色、对齐方式
- **高级样式**：边框、数字格式、文本换行
- **数据表格**：完整的员工信息表格示例
- **读取验证**：验证写入和读取的数据一致性
- **编辑功能**：演示文件编辑和更新
- **性能测试**：大数据量格式化性能测试

生成的文件包括：
- `comprehensive_formatting_test.xlsx` - 主测试文件
- `comprehensive_formatting_test_edited.xlsx` - 编辑测试文件
- `performance_test.xlsx` - 性能测试文件

## 📊 性能基准

在现代硬件上的性能表现：

| 操作 | 数据量 | 时间 | 内存使用 |
|------|--------|------|----------|
| 写入 | 100万单元格 | ~2.5秒 | ~150MB |
| 读取 | 100万单元格 | ~1.8秒 | ~120MB |
| 搜索替换 | 10万单元格 | ~0.3秒 | ~50MB |

## 🤝 贡献指南

我们欢迎各种形式的贡献！

### 如何贡献

1. Fork 项目
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启 Pull Request

### 代码规范

- 遵循 C++17 标准
- 使用一致的命名约定
- 添加适当的注释和文档
- 确保所有测试通过

## 📄 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情。

## 🆚 与其他库的比较

| 特性 | FastExcel | libxlsxwriter | OpenXLSX | xlsxio |
|------|-----------|---------------|----------|--------|
| 读取支持 | ✅ | ❌ | ✅ | ✅ |
| 写入支持 | ✅ | ✅ | ✅ | ✅ |
| 编辑支持 | ✅ | ❌ | ✅ | ❌ |
| 样式支持 | ✅ | ✅ | ✅ | ❌ |
| 内存优化 | ✅ | ✅ | ❌ | ✅ |
| 异常安全 | ✅ | ❌ | ✅ | ❌ |
| C++17 | ✅ | ❌ | ✅ | ❌ |

## 🔗 相关链接

- [完整 API 文档](docs/FastExcel_API_Documentation.md)
- [示例代码](examples/)
- [代码优化和测试总结](docs/FastExcel_Optimization_and_Testing_Summary.md)
- [设计模式改进建议](docs/Design_Pattern_Improvements.md)
- [性能分析报告](docs/FastExcel_ZIP_Performance_Analysis.md)
- [架构设计文档](docs/Cell_Optimization_Summary.md)

## 📞 支持和反馈

- 🐛 [报告 Bug](https://github.com/your-repo/FastExcel/issues)
- 💡 [功能请求](https://github.com/your-repo/FastExcel/issues)
- 📧 邮箱支持: support@fastexcel.com
- 💬 讨论区: [GitHub Discussions](https://github.com/your-repo/FastExcel/discussions)

## 🙏 致谢

感谢所有为 FastExcel 项目做出贡献的开发者和用户！

特别感谢：
- libxlsxwriter 项目提供的灵感
- 所有测试用户提供的宝贵反馈
- 开源社区的支持

---

**FastExcel** - 让 Excel 文件处理变得简单高效！ 🚀