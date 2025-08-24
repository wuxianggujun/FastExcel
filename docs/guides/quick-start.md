# FastExcel 快速开始指南

本指南将帮助你在5分钟内开始使用 FastExcel。

## 📋 前置要求

- C++17 兼容编译器 (GCC 7+, Clang 5+, MSVC 2017+)
- CMake 3.15+
- Git (用于克隆仓库)

## 🚀 安装步骤

### 1. 克隆项目
```bash
git clone https://github.com/wuxianggujun/FastExcel.git
cd FastExcel
```

### 2. 配置构建
```bash
cmake -B cmake-build-debug -S .
```

### 3. 编译项目
```bash
cmake --build cmake-build-debug -j 4
```

## 💡 第一个程序

### 创建简单的Excel文件

```cpp
#include "fastexcel/FastExcel.hpp"

using namespace fastexcel;
using namespace fastexcel::core;

int main() {
    try {
        // 创建工作簿
        auto workbook = std::make_unique<Workbook>("hello.xlsx");
        
        // 添加工作表
        auto worksheet = workbook->addWorksheet("Hello");
        
        // 设置单元格值
        worksheet->getCell("A1")->setValue("Hello, FastExcel!");
        worksheet->getCell("A2")->setValue(42.0);
        worksheet->getCell("A3")->setValue(true);
        
        // 保存文件
        workbook->save();
        
        std::cout << "Excel file created successfully!" << std::endl;
        
    } catch (const std::exception& e) {
        std::cerr << "Error: " << e.what() << std::endl;
        return 1;
    }
    
    return 0;
}
```

### 编译和运行
```bash
# 如果使用CMake构建的项目
cmake --build cmake-build-debug --target hello

# 或者直接编译
g++ -std=c++17 -I./src hello.cpp -L./cmake-build-debug/lib -lfastexcel -o hello
./hello
```

## 📊 数据类型支持

FastExcel 支持所有主要的Excel数据类型：

```cpp
// 数字
cell->setValue(123.45);
cell->setValue(42);

// 字符串 (自动优化短字符串)
cell->setValue("Hello World");
cell->setValue(std::string("Long text..."));

// 布尔值
cell->setValue(true);
cell->setValue(false);

// 公式
cell->setFormula("=SUM(A1:A10)", 100.0); // 结果可选

// 日期时间 (作为数字存储)
cell->setValue(44562.0); // Excel日期序列号
```

## 🎨 样式和格式化

使用现代的StyleBuilder API创建样式：

```cpp
// 创建样式
auto titleStyle = workbook->createStyleBuilder()
    .font().name("Arial").size(14).bold(true).color(Color::BLUE)
    .fill().pattern(PatternType::Solid).fgColor(Color::LIGHT_GRAY)
    .border().all(BorderStyle::Thin, Color::BLACK)
    .alignment().horizontal(HorizontalAlign::Center)
    .build();

// 应用样式
auto cell = worksheet->getCell("A1");
cell->setFormat(titleStyle);
```

## 📈 性能优化技巧

### 1. 启用高性能模式
```cpp
workbook->setHighPerformanceMode(true);
```

### 2. 批量操作
```cpp
// 批量设置数据
for (int row = 0; row < 1000; ++row) {
    for (int col = 0; col < 10; ++col) {
        worksheet->getCell(row, col)->setValue(row * col);
    }
}
```

### 3. 使用常量内存模式
```cpp
workbook->setConstantMemoryMode(true); // 大文件处理
```

## 🔍 错误处理

FastExcel 使用现代C++异常处理：

```cpp
try {
    auto workbook = Workbook::create("output.xlsx");
    auto worksheet = workbook->addWorksheet("Data");
    
    // 你的代码...
    
    workbook->save();
} catch (const std::exception& e) {
    std::cerr << "错误: " << e.what() << std::endl;
}
```

## 📊 读取Excel文件

```cpp
// 打开现有文件
auto workbook = Workbook::openForReading("input.xlsx");

if (workbook) {
    auto worksheet = workbook->getWorksheet(0); // 第一个工作表
    
    // 获取数据范围
    auto [maxRow, maxCol] = worksheet->getUsedRange();
    
    // 读取数据
    for (int row = 0; row <= maxRow; ++row) {
        for (int col = 0; col <= maxCol; ++col) {
            auto cell = worksheet->getCell(row, col);
            if (!cell->isEmpty()) {
                std::cout << "(" << row << "," << col << "): " 
                         << cell->getValue<std::string>() << std::endl;
            }
        }
    }
}
```

## 🖼️ 图片支持

```cpp
// 插入图片
auto image = std::make_unique<Image>("logo.png");
image->setPosition(100, 100);  // 像素位置
image->setSize(200, 150);      // 像素大小
worksheet->insertImage(std::move(image), AnchorType::Absolute);

// 单元格锚定图片
auto cellImage = std::make_unique<Image>("chart.jpg");
worksheet->insertImage(std::move(cellImage), AnchorType::OneCell, 2, 1); // C2
```

## 📄 CSV 集成

```cpp
// 从CSV创建Excel
CSVOptions options = CSVOptions::standard();
options.has_header = true;
options.auto_detect_delimiter = true;

auto worksheet = workbook->loadCSV("data.csv", "导入数据", options);

// 导出为CSV
worksheet->saveAsCSV("export.csv", CSVOptions::europeanStyle());
```

## 🧪 测试你的安装

运行示例程序：

```bash
cd cmake-build-debug/examples
./basic_usage_example
./formatting_example
./image_insertion_demo
```

## 📚 下一步

- 查看 [完整API文档](../api/) 了解所有功能
- 阅读 [架构设计文档](../architecture/) 理解内部实现
- 探索 [示例代码](../examples/) 学习最佳实践
- 查看 [性能优化指南](performance-guide.md) 提升应用性能

## ❓ 常见问题

### Q: 编译时出现找不到头文件的错误？
A: 确保包含路径正确，使用 `-I./src` 参数。

### Q: 链接时出现未定义符号错误？
A: 确保链接了正确的库文件，使用 `-L./cmake-build-debug/lib -lfastexcel`。

### Q: 如何处理中文文本？
A: FastExcel 完全支持 UTF-8 编码，直接使用中文字符串即可。

### Q: 如何处理大文件？
A: 使用 `setConstantMemoryMode(true)` 启用常量内存模式。

### Q: 可以在多线程环境中使用吗？
A: 是的，FastExcel 的核心组件是线程安全的，详见 [线程安全指南](thread-safety.md)。

---

**恭喜！** 🎉 你已经掌握了 FastExcel 的基本使用。现在可以开始构建你的Excel应用了！