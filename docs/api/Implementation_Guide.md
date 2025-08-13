# FastExcel 便捷 API 实现示例

## 概述

本文档提供具体的代码实现示例，展示如何为 FastExcel 添加更简洁的 API。

## 1. 核心模板函数实现

### 1.1 在 Cell 类中添加模板方法

```cpp
// 在 src/fastexcel/core/Cell.hpp 中添加
template<typename T>
T getValue() const {
    if constexpr (std::is_same_v<T, std::string>) {
        return getStringValue();
    } else if constexpr (std::is_floating_point_v<T>) {
        return static_cast<T>(getNumberValue());
    } else if constexpr (std::is_integral_v<T> && !std::is_same_v<T, bool>) {
        return static_cast<T>(getNumberValue());
    } else if constexpr (std::is_same_v<T, bool>) {
        return getBooleanValue();
    } else {
        static_assert(std::is_same_v<T, std::string>, 
                      "Unsupported type for Cell::getValue<T>()");
    }
}

template<typename T>
std::optional<T> tryGetValue() const noexcept {
    try {
        return getValue<T>();
    } catch (...) {
        return std::nullopt;
    }
}

template<typename T>
T getValueOr(const T& default_value) const noexcept {
    return tryGetValue<T>().value_or(default_value);
}

// 模板化设置方法
template<typename T>
void setValue(const T& value) {
    if constexpr (std::is_arithmetic_v<T> && !std::is_same_v<T, bool>) {
        setValue(static_cast<double>(value));
    } else if constexpr (std::is_same_v<T, bool>) {
        setValue(value);
    } else if constexpr (std::is_convertible_v<T, std::string>) {
        setValue(std::string(value));
    } else {
        static_assert(std::is_arithmetic_v<T>, 
                      "Unsupported type for Cell::setValue<T>()");
    }
}
```

### 1.2 在 Workbook 类中添加便捷方法

```cpp
// 在 src/fastexcel/core/Workbook.hpp 中添加
public:
    // 🚀 便捷访问：直接获取单元格值
    template<typename T>
    T getValue(size_t sheet_idx, int row, int col) const {
        ensureReadable("getValue");
        
        if (sheet_idx >= worksheets_.size()) {
            throw Exception("Sheet index out of range");
        }
        
        auto& cell = worksheets_[sheet_idx]->getCell(row, col);
        return cell.getValue<T>();
    }
    
    template<typename T>
    T getValue(const std::string& sheet_name, int row, int col) const {
        auto sheet = getWorksheet(sheet_name);
        if (!sheet) {
            throw Exception("Sheet not found: " + sheet_name);
        }
        
        auto& cell = sheet->getCell(row, col);
        return cell.getValue<T>();
    }
    
    template<typename T>
    T getValue(const std::string& address) const {
        // 解析地址格式 "Sheet1!A1" 或 "A1"
        auto [sheet_name, row, col] = parseAddress(address);
        
        if (sheet_name.empty()) {
            // 默认使用第一个工作表
            return getValue<T>(0, row, col);
        } else {
            return getValue<T>(sheet_name, row, col);
        }
    }
    
    // 🚀 便捷访问：直接设置单元格值
    template<typename T>
    void setValue(size_t sheet_idx, int row, int col, const T& value) {
        ensureEditable("setValue");
        
        if (sheet_idx >= worksheets_.size()) {
            throw Exception("Sheet index out of range");
        }
        
        worksheets_[sheet_idx]->getCell(row, col).setValue<T>(value);
        // 标记工作表为已修改
        if (dirty_manager_) {
            dirty_manager_->markSheetModified(sheet_idx);
        }
    }
    
    template<typename T>
    void setValue(const std::string& sheet_name, int row, int col, const T& value) {
        auto sheet = getWorksheet(sheet_name);
        if (!sheet) {
            throw Exception("Sheet not found: " + sheet_name);
        }
        
        sheet->getCell(row, col).setValue<T>(value);
        // 查找工作表索引并标记修改
        for (size_t i = 0; i < worksheets_.size(); ++i) {
            if (worksheets_[i].get() == sheet.get()) {
                if (dirty_manager_) {
                    dirty_manager_->markSheetModified(i);
                }
                break;
            }
        }
    }
    
    template<typename T>
    void setValue(const std::string& address, const T& value) {
        auto [sheet_name, row, col] = parseAddress(address);
        
        if (sheet_name.empty()) {
            setValue<T>(0, row, col, value);
        } else {
            setValue<T>(sheet_name, row, col, value);
        }
    }
    
    // 🚀 批量操作：获取范围
    template<typename T>
    std::vector<T> getRange(const std::string& range) const {
        auto [sheet_name, start_row, start_col, end_row, end_col] = parseRange(range);
        
        std::vector<T> result;
        result.reserve((end_row - start_row + 1) * (end_col - start_col + 1));
        
        auto sheet = sheet_name.empty() ? getWorksheet(0) : getWorksheet(sheet_name);
        if (!sheet) {
            throw Exception("Sheet not found: " + sheet_name);
        }
        
        for (int row = start_row; row <= end_row; ++row) {
            for (int col = start_col; col <= end_col; ++col) {
                result.push_back(sheet->getCell(row, col).getValue<T>());
            }
        }
        
        return result;
    }
    
    // 🚀 批量操作：设置范围
    template<typename T>
    void setRange(const std::string& range, const std::vector<T>& values) {
        auto [sheet_name, start_row, start_col, end_row, end_col] = parseRange(range);
        
        auto sheet = sheet_name.empty() ? getWorksheet(0) : getWorksheet(sheet_name);
        if (!sheet) {
            throw Exception("Sheet not found: " + sheet_name);
        }
        
        size_t value_idx = 0;
        for (int row = start_row; row <= end_row && value_idx < values.size(); ++row) {
            for (int col = start_col; col <= end_col && value_idx < values.size(); ++col) {
                sheet->getCell(row, col).setValue<T>(values[value_idx++]);
            }
        }
        
        // 标记修改
        for (size_t i = 0; i < worksheets_.size(); ++i) {
            if (worksheets_[i].get() == sheet.get()) {
                if (dirty_manager_) {
                    dirty_manager_->markSheetModified(i);
                }
                break;
            }
        }
    }

private:
    // 地址解析辅助函数
    std::tuple<std::string, int, int> parseAddress(const std::string& address) const;
    std::tuple<std::string, int, int, int, int> parseRange(const std::string& range) const;
```

## 2. 地址解析实现

### 2.1 实现文件 `src/fastexcel/core/AddressParser.hpp`

```cpp
#pragma once
#include <string>
#include <tuple>
#include <regex>

namespace fastexcel {
namespace core {

class AddressParser {
public:
    // 解析单个地址 "Sheet1!A1" 或 "A1"
    // 返回 (sheet_name, row, col)，row和col基于0
    static std::tuple<std::string, int, int> parseAddress(const std::string& address) {
        // 正则表达式匹配 [SheetName!]A1 格式
        static const std::regex addr_regex(R"(^(?:([^!]+)!)?([A-Z]+)(\d+)$)");
        std::smatch matches;
        
        if (!std::regex_match(address, matches, addr_regex)) {
            throw Exception("Invalid address format: " + address);
        }
        
        std::string sheet_name = matches[1].str();
        std::string col_str = matches[2].str();
        int row = std::stoi(matches[3].str()) - 1; // 转换为0基索引
        
        int col = columnStringToIndex(col_str);
        
        return std::make_tuple(sheet_name, row, col);
    }
    
    // 解析范围 "Sheet1!A1:C3" 或 "A1:C3"
    // 返回 (sheet_name, start_row, start_col, end_row, end_col)
    static std::tuple<std::string, int, int, int, int> parseRange(const std::string& range) {
        // 查找冒号分隔符
        auto colon_pos = range.find(':');
        if (colon_pos == std::string::npos) {
            // 单个单元格当作1x1范围
            auto [sheet, row, col] = parseAddress(range);
            return std::make_tuple(sheet, row, col, row, col);
        }
        
        std::string start_addr = range.substr(0, colon_pos);
        std::string end_addr = range.substr(colon_pos + 1);
        
        auto [start_sheet, start_row, start_col] = parseAddress(start_addr);
        auto [end_sheet, end_row, end_col] = parseAddress(end_addr);
        
        // 如果结束地址没有指定工作表，使用开始地址的工作表
        if (end_sheet.empty()) {
            end_sheet = start_sheet;
        }
        
        // 工作表名必须相同
        if (start_sheet != end_sheet) {
            throw Exception("Range cannot span multiple sheets");
        }
        
        return std::make_tuple(start_sheet, start_row, start_col, end_row, end_col);
    }

private:
    // 将列字符串转换为索引 (A->0, B->1, ..., Z->25, AA->26, ...)
    static int columnStringToIndex(const std::string& col_str) {
        int result = 0;
        for (char c : col_str) {
            result = result * 26 + (c - 'A' + 1);
        }
        return result - 1; // 转换为0基索引
    }
};

}} // namespace fastexcel::core
```

### 2.2 在 Workbook.cpp 中实现解析方法

```cpp
// 在 src/fastexcel/core/Workbook.cpp 中添加
std::tuple<std::string, int, int> Workbook::parseAddress(const std::string& address) const {
    return AddressParser::parseAddress(address);
}

std::tuple<std::string, int, int, int, int> Workbook::parseRange(const std::string& range) const {
    return AddressParser::parseRange(range);
}
```

## 3. 工作表代理类实现

### 3.1 WorksheetProxy 类定义

```cpp
// 在 src/fastexcel/core/WorksheetProxy.hpp 中创建
#pragma once
#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/Workbook.hpp"
#include <memory>

namespace fastexcel {
namespace core {

class WorksheetProxy {
private:
    std::shared_ptr<Worksheet> worksheet_;
    std::weak_ptr<Workbook> workbook_;
    
public:
    WorksheetProxy(std::shared_ptr<Worksheet> worksheet, std::shared_ptr<Workbook> workbook)
        : worksheet_(std::move(worksheet)), workbook_(workbook) {}
    
    // 🚀 链式调用设置值
    template<typename T>
    WorksheetProxy& set(int row, int col, const T& value) {
        worksheet_->getCell(row, col).setValue<T>(value);
        
        // 标记修改状态
        if (auto wb = workbook_.lock()) {
            // 查找工作表索引
            for (size_t i = 0; i < wb->getWorksheetCount(); ++i) {
                if (wb->getWorksheet(i).get() == worksheet_.get()) {
                    if (auto dirty_mgr = wb->getDirtyManager()) {
                        dirty_mgr->markSheetModified(i);
                    }
                    break;
                }
            }
        }
        
        return *this;
    }
    
    // 🚀 获取值
    template<typename T>
    T get(int row, int col) const {
        return worksheet_->getCell(row, col).getValue<T>();
    }
    
    // 🚀 范围操作
    template<typename T>
    WorksheetProxy& setRange(const std::string& range, const std::vector<T>& values) {
        auto [_, start_row, start_col, end_row, end_col] = AddressParser::parseRange(range);
        
        size_t value_idx = 0;
        for (int row = start_row; row <= end_row && value_idx < values.size(); ++row) {
            for (int col = start_col; col <= end_col && value_idx < values.size(); ++col) {
                worksheet_->getCell(row, col).setValue<T>(values[value_idx++]);
            }
        }
        
        return *this;
    }
    
    template<typename T>
    std::vector<T> getRange(const std::string& range) const {
        auto [_, start_row, start_col, end_row, end_col] = AddressParser::parseRange(range);
        
        std::vector<T> result;
        result.reserve((end_row - start_row + 1) * (end_col - start_col + 1));
        
        for (int row = start_row; row <= end_row; ++row) {
            for (int col = start_col; col <= end_col; ++col) {
                result.push_back(worksheet_->getCell(row, col).getValue<T>());
            }
        }
        
        return result;
    }
    
    // 🚀 格式设置
    WorksheetProxy& format(int row, int col, const FormatDescriptor& fmt) {
        worksheet_->getCell(row, col).setFormat(std::make_shared<FormatDescriptor>(fmt));
        return *this;
    }
    
    // 直接访问原始工作表
    Worksheet* operator->() { return worksheet_.get(); }
    const Worksheet* operator->() const { return worksheet_.get(); }
    Worksheet& operator*() { return *worksheet_; }
    const Worksheet& operator*() const { return *worksheet_; }
};

}} // namespace fastexcel::core
```

### 3.2 在 Workbook 中添加代理方法

```cpp
// 在 Workbook.hpp 中添加
#include "fastexcel/core/WorksheetProxy.hpp"

public:
    // 🚀 返回工作表代理
    WorksheetProxy sheet(size_t index) {
        if (index >= worksheets_.size()) {
            throw Exception("Sheet index out of range");
        }
        return WorksheetProxy(worksheets_[index], shared_from_this());
    }
    
    WorksheetProxy sheet(const std::string& name) {
        auto ws = getWorksheet(name);
        if (!ws) {
            throw Exception("Sheet not found: " + name);
        }
        return WorksheetProxy(ws, shared_from_this());
    }
    
    // 🚀 操作符重载
    WorksheetProxy operator[](size_t index) { 
        return sheet(index); 
    }
    
    WorksheetProxy operator[](const std::string& name) { 
        return sheet(name); 
    }
```

## 4. 使用示例和测试

### 4.1 基本使用示例

```cpp
#include "fastexcel/FastExcel.hpp"

void example_new_api() {
    // 🚀 创建工作簿并使用新 API
    auto wb = fastexcel::Workbook::create("test.xlsx");
    
    // 直接设置值 - 简洁API
    wb->setValue(0, 0, 0, "产品名称");        // Sheet0, A1
    wb->setValue(0, 0, 1, "价格");           // Sheet0, B1
    wb->setValue(0, 1, 0, "iPhone 15");      // Sheet0, A2
    wb->setValue(0, 1, 1, 999.99);          // Sheet0, B2
    
    // 批量设置
    wb->setRange("Sheet1!C1:C5", std::vector<int>{1, 2, 3, 4, 5});
    
    // 链式调用
    wb->sheet(0)
      .set(2, 0, "iPad Pro")
      .set(2, 1, 1299.99)
      .set(3, 0, "MacBook")
      .set(3, 1, 2399.99);
    
    // 类型安全获取
    std::string product = wb->getValue<std::string>(0, 1, 0);
    double price = wb->getValue<double>(0, 1, 1);
    
    // 安全获取（避免异常）
    auto cell = wb->sheet(0)->getCell(10, 10);
    if (auto value = cell.tryGetValue<double>()) {
        std::cout << "数值: " << *value << "\n";
    } else {
        std::cout << "空单元格或非数值\n";
    }
    
    // 带默认值获取
    std::string name = cell.getValueOr<std::string>("未知产品");
    
    wb->save();
}
```

### 4.2 性能测试示例

```cpp
void performance_test() {
    auto wb = fastexcel::Workbook::create("perf_test.xlsx");
    
    // 🚀 批量操作性能测试
    const int ROWS = 10000;
    const int COLS = 10;
    
    auto start = std::chrono::high_resolution_clock::now();
    
    // 旧方式（逐个设置）
    for (int row = 0; row < ROWS; ++row) {
        for (int col = 0; col < COLS; ++col) {
            wb->getWorksheet(0)->getCell(row, col).setValue(row * col);
        }
    }
    
    auto mid = std::chrono::high_resolution_clock::now();
    
    // 新方式（批量设置）
    std::vector<double> values;
    values.reserve(ROWS * COLS);
    for (int row = 0; row < ROWS; ++row) {
        for (int col = 0; col < COLS; ++col) {
            values.push_back(row * col);
        }
    }
    wb->setRange("Sheet2!A1:J10000", values);
    
    auto end = std::chrono::high_resolution_clock::now();
    
    auto old_time = std::chrono::duration_cast<std::chrono::milliseconds>(mid - start);
    auto new_time = std::chrono::duration_cast<std::chrono::milliseconds>(end - mid);
    
    std::cout << "旧API时间: " << old_time.count() << "ms\n";
    std::cout << "新API时间: " << new_time.count() << "ms\n";
    std::cout << "性能提升: " << (double)old_time.count() / new_time.count() << "x\n";
}
```

### 4.3 单元测试

```cpp
#include <gtest/gtest.h>
#include "fastexcel/FastExcel.hpp"

class NewAPITest : public ::testing::Test {
protected:
    void SetUp() override {
        wb = fastexcel::Workbook::create("test_new_api.xlsx");
    }
    
    std::unique_ptr<fastexcel::Workbook> wb;
};

TEST_F(NewAPITest, DirectAccess) {
    // 测试直接访问API
    wb->setValue(0, 0, 0, "测试");
    wb->setValue(0, 0, 1, 123.45);
    wb->setValue(0, 0, 2, true);
    
    EXPECT_EQ(wb->getValue<std::string>(0, 0, 0), "测试");
    EXPECT_DOUBLE_EQ(wb->getValue<double>(0, 0, 1), 123.45);
    EXPECT_EQ(wb->getValue<bool>(0, 0, 2), true);
}

TEST_F(NewAPITest, AddressParsing) {
    // 测试地址解析
    wb->setValue("A1", "第一个");
    wb->setValue("Sheet1!B2", 42.0);
    
    EXPECT_EQ(wb->getValue<std::string>("A1"), "第一个");
    EXPECT_DOUBLE_EQ(wb->getValue<double>("Sheet1!B2"), 42.0);
}

TEST_F(NewAPITest, BatchOperations) {
    // 测试批量操作
    std::vector<int> values = {1, 2, 3, 4, 5, 6};
    wb->setRange("A1:C2", values);
    
    auto result = wb->getRange<int>("A1:C2");
    EXPECT_EQ(result, values);
}

TEST_F(NewAPITest, ChainedCalls) {
    // 测试链式调用
    wb->sheet(0)
      .set(0, 0, "产品")
      .set(0, 1, "价格")
      .set(1, 0, "iPhone")
      .set(1, 1, 999.99);
    
    EXPECT_EQ(wb->getValue<std::string>(0, 0, 0), "产品");
    EXPECT_EQ(wb->getValue<std::string>(0, 1, 0), "iPhone");
    EXPECT_DOUBLE_EQ(wb->getValue<double>(0, 1, 1), 999.99);
}

TEST_F(NewAPITest, TypeSafety) {
    // 测试类型安全
    wb->setValue(0, 0, 0, "非数字");
    
    auto cell = wb->sheet(0)->getCell(0, 0);
    EXPECT_FALSE(cell.tryGetValue<double>().has_value());
    EXPECT_EQ(cell.getValueOr<double>(0.0), 0.0);
    EXPECT_EQ(cell.getValueOr<std::string>("默认"), "非数字");
}
```

## 5. 向后兼容性

### 5.1 渐进式迁移

```cpp
// 在现有类中添加别名，保持兼容性
class Cell {
public:
    // 新 API
    template<typename T>
    T getValue() const { /* 新实现 */ }
    
    // 保持旧 API，但标记为弃用
    [[deprecated("Use getValue<std::string>() instead")]]
    std::string getStringValue() const { return getValue<std::string>(); }
    
    [[deprecated("Use getValue<double>() instead")]]
    double getNumberValue() const { return getValue<double>(); }
    
    [[deprecated("Use getValue<bool>() instead")]]
    bool getBooleanValue() const { return getValue<bool>(); }
};
```

### 5.2 迁移助手

```cpp
// 创建迁移助手宏
#define FASTEXCEL_OLD_API_WARNING \
    static_assert(false, "This API is deprecated. Please use the new template-based API.")

// 条件编译支持
#ifdef FASTEXCEL_ENABLE_OLD_API
    // 保留旧 API
#else
    // 仅新 API
    FASTEXCEL_OLD_API_WARNING
#endif
```

## 总结

这套新 API 实现提供了：

1. **类型安全**：编译期类型检查，运行时安全获取
2. **简洁语法**：一行代码完成复杂操作
3. **高性能**：批量操作，减少函数调用开销
4. **现代 C++**：利用 constexpr、模板特化等特性
5. **向后兼容**：渐进式迁移，不破坏现有代码

建议按照以下顺序实现：

1. **第一步**：实现 `Cell::getValue<T>()` 模板方法
2. **第二步**：添加 `Workbook` 的直接访问方法
3. **第三步**：实现地址解析和批量操作
4. **第四步**：添加工作表代理类和链式调用
5. **第五步**：完善测试和文档

这将显著提升 FastExcel 的易用性和性能。

---

*实现指南版本*：v1.0  
*最后更新*：2025-01-11