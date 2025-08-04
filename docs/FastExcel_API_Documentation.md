# FastExcel API 文档

## 概述

FastExcel 是一个高性能的 C++ Excel 文件处理库，支持读取、写入和编辑 XLSX 文件。本库设计注重性能优化、内存管理和异常安全。

## 核心特性

- **高性能读写**: 优化的内存管理和缓存系统
- **完整的编辑功能**: 支持单元格、行、列和工作表级别的编辑操作
- **样式支持**: 完整的 Excel 样式解析和应用
- **异常安全**: 完善的错误处理和异常管理系统
- **内存优化**: 内存池和 LRU 缓存系统
- **线程安全**: 关键操作的线程安全保证

## 快速开始

### 基本用法

```cpp
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/reader/XLSXReader.hpp"

using namespace fastexcel;

// 创建新的工作簿
auto workbook = core::Workbook::create("example.xlsx");
workbook->open();

// 添加工作表
auto worksheet = workbook->addWorksheet("Sheet1");

// 写入数据
worksheet->writeString(0, 0, "Hello");
worksheet->writeNumber(0, 1, 123.45);
worksheet->writeBoolean(0, 2, true);

// 保存文件
workbook->save();
workbook->close();

// 读取现有文件
reader::XLSXReader reader("example.xlsx");
reader.open();
auto loaded_workbook = reader.loadWorkbook();
reader.close();
```

## 核心类参考

### Workbook 类

工作簿是 Excel 文件的主要容器，包含一个或多个工作表。

#### 构造和生命周期

```cpp
// 创建新工作簿
static std::unique_ptr<Workbook> create(const std::string& filename);

// 加载现有工作簿进行编辑
static std::unique_ptr<Workbook> loadForEdit(const std::string& filename);

// 打开工作簿
bool open();

// 保存工作簿
bool save();
bool saveAs(const std::string& filename);

// 关闭工作簿
bool close();
```

#### 工作表管理

```cpp
// 添加工作表
std::shared_ptr<Worksheet> addWorksheet(const std::string& name);

// 获取工作表
std::shared_ptr<Worksheet> getWorksheet(const std::string& name);
std::shared_ptr<Worksheet> getWorksheet(size_t index);

// 删除工作表
bool removeWorksheet(const std::string& name);

// 重命名工作表
bool renameWorksheet(const std::string& old_name, const std::string& new_name);

// 获取工作表数量
size_t getWorksheetCount() const;

// 获取所有工作表名称
std::vector<std::string> getWorksheetNames() const;
```

#### 批量操作

```cpp
// 批量重命名工作表
int batchRenameWorksheets(const std::unordered_map<std::string, std::string>& rename_map);

// 批量删除工作表
int batchRemoveWorksheets(const std::vector<std::string>& names);

// 重新排序工作表
bool reorderWorksheets(const std::vector<std::string>& new_order);
```

#### 全局搜索和替换

```cpp
struct FindReplaceOptions {
    bool case_sensitive = false;
    bool whole_word = false;
    bool use_regex = false;
    std::vector<std::string> worksheet_filter; // 空表示所有工作表
};

// 全局查找
std::vector<CellLocation> findAll(const std::string& search_text, 
                                 const FindReplaceOptions& options = {});

// 全局替换
int findAndReplaceAll(const std::string& search_text, 
                     const std::string& replace_text,
                     const FindReplaceOptions& options = {});
```

### Worksheet 类

工作表包含单元格数据和格式信息。

#### 基本数据操作

```cpp
// 写入数据
void writeString(int row, int col, const std::string& value);
void writeNumber(int row, int col, double value);
void writeBoolean(int row, int col, bool value);
void writeFormula(int row, int col, const std::string& formula);

// 读取数据
const Cell& getCell(int row, int col) const;
std::string getCellString(int row, int col) const;
double getCellNumber(int row, int col) const;
bool getCellBoolean(int row, int col) const;
```

#### 单元格编辑

```cpp
// 编辑单元格值
bool editCellValue(int row, int col, const std::string& new_value, bool preserve_format = true);
bool editCellValue(int row, int col, double new_value, bool preserve_format = true);
bool editCellValue(int row, int col, bool new_value, bool preserve_format = true);

// 复制单元格
bool copyCell(int src_row, int src_col, int dst_row, int dst_col, bool copy_format = true);

// 移动单元格
bool moveCell(int src_row, int src_col, int dst_row, int dst_col);
```

#### 范围操作

```cpp
// 复制范围
bool copyRange(int src_start_row, int src_start_col, int src_end_row, int src_end_col,
               int dst_start_row, int dst_start_col, bool copy_format = true);

// 移动范围
bool moveRange(int src_start_row, int src_start_col, int src_end_row, int src_end_col,
               int dst_start_row, int dst_start_col);

// 删除范围
bool deleteRange(int start_row, int start_col, int end_row, int end_col);

// 插入行/列
bool insertRows(int start_row, int count);
bool insertColumns(int start_col, int count);

// 删除行/列
bool deleteRows(int start_row, int count);
bool deleteColumns(int start_col, int count);
```

#### 搜索和替换

```cpp
struct CellLocation {
    int row;
    int col;
    std::string value;
};

// 查找单元格
std::vector<CellLocation> findCells(const std::string& search_text, 
                                   bool case_sensitive = false, 
                                   bool whole_word = false);

// 查找和替换
int findAndReplace(const std::string& search_text, 
                  const std::string& replace_text,
                  bool case_sensitive = false, 
                  bool whole_word = false);
```

#### 排序

```cpp
// 排序范围
bool sortRange(int start_row, int start_col, int end_row, int end_col,
               int sort_column, bool ascending = true, bool has_header = false);
```

### XLSXReader 类

用于读取现有的 XLSX 文件。

```cpp
// 构造函数
explicit XLSXReader(const std::string& filename);

// 打开文件
bool open();

// 加载整个工作簿
std::unique_ptr<core::Workbook> loadWorkbook();

// 加载单个工作表
std::shared_ptr<core::Worksheet> loadWorksheet(const std::string& name);

// 获取工作表名称列表
std::vector<std::string> getWorksheetNames();

// 获取元数据
WorkbookMetadata getMetadata();

// 关闭文件
bool close();
```

### Format 类

用于设置单元格格式。

```cpp
// 字体设置
void setFontName(const std::string& name);
void setFontSize(double size);
void setFontColor(const Color& color);
void setBold(bool bold);
void setItalic(bool italic);
void setUnderline(bool underline);

// 填充设置
void setBackgroundColor(const Color& color);
void setPatternType(PatternType type);

// 边框设置
void setBorder(BorderPosition position, BorderStyle style, const Color& color);

// 对齐设置
void setHorizontalAlignment(HorizontalAlignment alignment);
void setVerticalAlignment(VerticalAlignment alignment);

// 数字格式
void setNumberFormat(const std::string& format);
```

## 性能优化

### 内存管理

FastExcel 提供了内存池和缓存系统来优化性能：

```cpp
#include "fastexcel/core/MemoryPool.hpp"
#include "fastexcel/core/CacheSystem.hpp"

// 获取内存管理器
auto& memory_manager = core::MemoryManager::getInstance();
auto& default_pool = memory_manager.getDefaultPool();

// 获取缓存管理器
auto& cache_manager = core::CacheManager::getInstance();
auto& string_cache = cache_manager.getStringCache();

// 获取统计信息
auto memory_stats = memory_manager.getGlobalStatistics();
auto cache_stats = cache_manager.getGlobalStatistics();
```

### 高性能模式

```cpp
// 启用高性能模式
workbook->setHighPerformanceMode(true);

// 设置缓冲区大小
workbook->setBufferSize(1024 * 1024); // 1MB

// 批量操作时禁用自动保存
workbook->setAutoSave(false);
// ... 执行批量操作 ...
workbook->save(); // 手动保存
```

## 异常处理

FastExcel 提供了完善的异常处理系统：

```cpp
#include "fastexcel/core/Exception.hpp"

try {
    auto workbook = core::Workbook::create("test.xlsx");
    workbook->open();
    // ... 操作 ...
} catch (const core::FileException& e) {
    std::cerr << "文件错误: " << e.getDetailedMessage() << std::endl;
    std::cerr << "文件名: " << e.getFilename() << std::endl;
} catch (const core::MemoryException& e) {
    std::cerr << "内存错误: " << e.getDetailedMessage() << std::endl;
    std::cerr << "请求大小: " << e.getRequestedSize() << std::endl;
} catch (const core::FastExcelException& e) {
    std::cerr << "FastExcel错误: " << e.getDetailedMessage() << std::endl;
}
```

### 自定义错误处理

```cpp
class CustomErrorHandler : public core::ErrorHandler {
public:
    bool handleError(const core::FastExcelException& exception) override {
        // 自定义错误处理逻辑
        logError(exception);
        return true; // 继续执行
    }
    
    void handleWarning(const std::string& message, const std::string& context) override {
        // 自定义警告处理逻辑
        logWarning(message, context);
    }
};

// 设置自定义错误处理器
core::ErrorManager::getInstance().setErrorHandler(
    std::make_unique<CustomErrorHandler>());
```

## 示例代码

### 完整的读写示例

```cpp
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/reader/XLSXReader.hpp"
#include "fastexcel/core/Exception.hpp"

int main() {
    try {
        // 创建新工作簿
        auto workbook = fastexcel::core::Workbook::create("example.xlsx");
        workbook->open();
        
        // 添加工作表
        auto sheet = workbook->addWorksheet("数据表");
        
        // 写入表头
        sheet->writeString(0, 0, "姓名");
        sheet->writeString(0, 1, "年龄");
        sheet->writeString(0, 2, "部门");
        
        // 写入数据
        std::vector<std::tuple<std::string, int, std::string>> data = {
            {"张三", 25, "技术部"},
            {"李四", 30, "销售部"},
            {"王五", 28, "人事部"}
        };
        
        for (size_t i = 0; i < data.size(); ++i) {
            sheet->writeString(i + 1, 0, std::get<0>(data[i]));
            sheet->writeNumber(i + 1, 1, std::get<1>(data[i]));
            sheet->writeString(i + 1, 2, std::get<2>(data[i]));
        }
        
        // 保存并关闭
        workbook->save();
        workbook->close();
        
        // 读取文件
        fastexcel::reader::XLSXReader reader("example.xlsx");
        reader.open();
        
        auto loaded_workbook = reader.loadWorkbook();
        auto loaded_sheet = loaded_workbook->getWorksheet("数据表");
        
        // 读取数据
        for (int row = 0; row < 4; ++row) {
            for (int col = 0; col < 3; ++col) {
                std::cout << loaded_sheet->getCellString(row, col) << "\t";
            }
            std::cout << std::endl;
        }
        
        reader.close();
        
    } catch (const fastexcel::core::FastExcelException& e) {
        std::cerr << "错误: " << e.getDetailedMessage() << std::endl;
        return 1;
    }
    
    return 0;
}
```

### 编辑现有文件

```cpp
#include "fastexcel/core/Workbook.hpp"

int main() {
    try {
        // 加载现有文件进行编辑
        auto workbook = fastexcel::core::Workbook::loadForEdit("existing.xlsx");
        auto sheet = workbook->getWorksheet("Sheet1");
        
        // 查找和替换
        int replacements = sheet->findAndReplace("旧值", "新值", false, false);
        std::cout << "替换了 " << replacements << " 个单元格" << std::endl;
        
        // 插入新行
        sheet->insertRows(1, 1);
        sheet->writeString(1, 0, "新插入的行");
        
        // 复制范围
        sheet->copyRange(0, 0, 0, 2, 10, 0, true);
        
        // 排序数据
        sheet->sortRange(1, 0, 10, 2, 1, true, true);
        
        // 保存更改
        workbook->save();
        
    } catch (const fastexcel::core::FastExcelException& e) {
        std::cerr << "错误: " << e.getDetailedMessage() << std::endl;
        return 1;
    }
    
    return 0;
}
```

## 编译和链接

### CMake 配置

```cmake
find_package(FastExcel REQUIRED)
target_link_libraries(your_target FastExcel::FastExcel)
```

### 手动编译

```bash
g++ -std=c++17 -I/path/to/fastexcel/include \
    your_code.cpp -L/path/to/fastexcel/lib -lfastexcel
```

## 性能建议

1. **批量操作**: 对于大量数据操作，使用批量API而不是逐个单元格操作
2. **内存管理**: 启用高性能模式和适当的缓冲区大小
3. **异常处理**: 在性能关键的代码中使用非抛出版本的API
4. **缓存**: 利用内置的缓存系统来提高重复访问的性能
5. **线程安全**: 在多线程环境中，为每个线程使用独立的工作簿实例

## 限制和注意事项

1. **文件格式**: 目前只支持 XLSX 格式，不支持旧的 XLS 格式
2. **公式**: 支持公式的读取和写入，但不支持公式计算
3. **图表**: 暂不支持图表的创建和编辑
4. **宏**: 不支持 VBA 宏
5. **内存使用**: 大文件处理时注意内存使用量

## 版本历史

- **v1.0.0**: 初始版本，基本读写功能
- **v1.1.0**: 添加编辑功能和样式支持
- **v1.2.0**: 性能优化和内存管理
- **v1.3.0**: 异常处理和错误管理系统

## 许可证

本项目采用 MIT 许可证，详见 LICENSE 文件。

## 贡献

欢迎提交 Issue 和 Pull Request。在提交代码前，请确保：

1. 代码符合项目的编码规范
2. 添加了适当的测试用例
3. 更新了相关文档

## 支持

如有问题或建议，请通过以下方式联系：

- GitHub Issues: [项目地址]
- 邮箱: [联系邮箱]
- 文档: [在线文档地址]