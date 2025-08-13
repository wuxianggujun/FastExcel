# FastExcel 便捷 API 快速参考

## 🚀 新 API vs 旧 API 对比

| 操作 | 旧 API | 新 API | 效率提升 |
|------|---------|---------|----------|
| **获取字符串值** | `wb->getWorksheet(0)->getCell(1,1).getStringValue()` | `wb->getValue<std::string>(0, 1, 1)` | 70%减少 |
| **设置数值** | `wb->getWorksheet(0)->getCell(1,1).setValue(99.9)` | `wb->setValue(0, 1, 1, 99.9)` | 65%减少 |
| **类型安全获取** | 需要先检查类型再获取 | `cell.tryGetValue<double>()` | 安全+简洁 |
| **批量设置** | 循环逐个设置 | `wb->setRange("A1:C3", values)` | 90%减少 |
| **地址访问** | 复杂的坐标计算 | `wb->getValue<int>("Sheet1!A1")` | 直观易用 |

## 📖 常用操作速查

### 1. 基本读写操作

```cpp
auto wb = Workbook::open("data.xlsx");

// ✅ 读取数据 - 自动类型转换
std::string name = wb->getValue<std::string>(0, 1, 1);   // Sheet0, B2
double price = wb->getValue<double>("产品表", 1, 2);      // 按表名访问  
int count = wb->getValue<int>("Sheet1!C2");              // Excel地址格式

// ✅ 写入数据 - 类型安全
wb->setValue(0, 1, 1, "新产品");                         // 字符串
wb->setValue(0, 1, 2, 199.99);                          // 数值
wb->setValue("Sheet1!A1", true);                        // 布尔值
```

### 2. 批量操作

```cpp
// ✅ 批量读取
auto prices = wb->getRange<double>("Sheet1!B1:B10");     // 获取价格列
auto names = wb->getRange<std::string>("A1:A5");         // 获取名称

// ✅ 批量写入  
wb->setRange("C1:C5", std::vector<int>{1,2,3,4,5});     // 设置序号
wb->setRange("Sheet2!A1:D1", 
    std::vector<std::string>{"姓名","年龄","城市","薪资"}); // 表头
```

### 3. 链式调用

```cpp
// ✅ 流式操作 - 优雅的数据设置
wb->sheet(0)
  .set(0, 0, "产品编码").set(0, 1, "产品名称").set(0, 2, "单价")
  .set(1, 0, "P001")    .set(1, 1, "iPhone 15") .set(1, 2, 999.99)
  .set(2, 0, "P002")    .set(2, 1, "iPad Pro")  .set(2, 2, 1299.99);

// ✅ 使用下标操作符  
wb[0].set(3, 0, "P003").set(3, 1, "MacBook").set(3, 2, 2399.99);
wb["销售表"].setRange("A1:C1", std::vector<std::string>{"日期","金额","备注"});
```

### 4. 类型安全访问

```cpp
auto cell = wb->sheet(0)->getCell(1, 1);

// ✅ 模板化获取 - 编译期类型检查
std::string text = cell.getValue<std::string>();
double number = cell.getValue<double>();  
int integer = cell.getValue<int>();

// ✅ 安全获取 - 避免异常
if (auto value = cell.tryGetValue<double>()) {
    std::cout << "数值: " << *value << std::endl;
} else {
    std::cout << "不是数值类型" << std::endl;
}

// ✅ 带默认值获取
std::string name = cell.getValueOr<std::string>("未知");
double price = cell.getValueOr<double>(0.0);
```

## 🎯 高效使用模式

### 1. 数据导入模式

```cpp
// 高效读取大量数据
auto wb = Workbook::openForReading("large_data.xlsx");  // 只读模式

// 批量读取，性能最优
auto data = wb->getRange<double>("Sheet1!A1:Z1000");   
for (size_t i = 0; i < data.size(); ++i) {
    processData(data[i]);  // 处理数据
}

// 流式处理超大文件
for (auto iter = wb->iterateCells(0); iter.hasNext();) {
    auto cell = iter.next();
    if (cell.type == CellType::Number) {
        processNumber(cell.value);
    }
}
```

### 2. 报表生成模式  

```cpp
auto wb = Workbook::create("report.xlsx");

// 快速设置表头
wb->setRange("A1:E1", std::vector<std::string>{
    "序号", "产品名称", "销量", "单价", "金额"
});

// 批量写入数据行
std::vector<ReportRow> data = fetchReportData();
for (size_t i = 0; i < data.size(); ++i) {
    int row = i + 2;  // 从第2行开始
    wb->setValue(0, row, 0, static_cast<int>(i + 1));     // 序号
    wb->setValue(0, row, 1, data[i].product_name);        // 产品名
    wb->setValue(0, row, 2, data[i].quantity);            // 销量  
    wb->setValue(0, row, 3, data[i].unit_price);          // 单价
    wb->setValue(0, row, 4, data[i].total_amount);        // 金额
}

// 或使用链式调用
for (size_t i = 0; i < data.size(); ++i) {
    int row = i + 2;
    wb->sheet(0)
      .set(row, 0, static_cast<int>(i + 1))
      .set(row, 1, data[i].product_name)
      .set(row, 2, data[i].quantity)
      .set(row, 3, data[i].unit_price)
      .set(row, 4, data[i].total_amount);
}
```

### 3. 数据分析模式

```cpp
auto wb = Workbook::openForReading("sales_data.xlsx");

// 快速统计分析
auto prices = wb->getRange<double>("B:B");              // 整列数据
auto quantities = wb->getRange<int>("C:C");

// 计算统计信息
double total_revenue = 0;
int total_quantity = 0;
for (size_t i = 0; i < prices.size(); ++i) {
    total_revenue += prices[i] * quantities[i];
    total_quantity += quantities[i];
}

double avg_price = total_revenue / total_quantity;
std::cout << "平均单价: " << avg_price << std::endl;
```

### 4. 配置文件模式

```cpp
// 将Excel作为配置文件使用
auto config = Workbook::openForReading("config.xlsx");

// 读取配置项
std::string server_host = config->getValue<std::string>("配置!B1");
int server_port = config->getValue<int>("配置!B2");  
bool debug_mode = config->getValue<bool>("配置!B3");
double timeout = config->getValue<double>("配置!B4");

// 安全读取（提供默认值）
auto cell_b5 = config->sheet("配置")->getCell(4, 1);
int retry_count = cell_b5.getValueOr<int>(3);           // 默认重试3次
std::string log_level = cell_b5.getValueOr<std::string>("INFO");  
```

## ⚡ 性能优化技巧

### 1. 选择合适的打开模式

```cpp
// 🚀 只读数据：使用只读模式，内存占用更少
auto wb = Workbook::openForReading("data.xlsx");

// 🚀 需要编辑：使用编辑模式
auto wb = Workbook::openForEditing("data.xlsx");

// 🚀 创建新文件：直接创建
auto wb = Workbook::create("new_file.xlsx");
```

### 2. 批量操作优于逐个操作

```cpp
// ❌ 低效：逐个设置
for (int i = 0; i < 1000; ++i) {
    wb->setValue(0, i, 0, values[i]);
}

// ✅ 高效：批量设置
wb->setRange("A1:A1000", values);
```

### 3. 重用工作表引用

```cpp
// ❌ 低效：重复获取工作表
for (int i = 0; i < 100; ++i) {
    wb->setValue("数据表", i, 0, data[i]);  // 每次都查找工作表
}

// ✅ 高效：重用工作表代理
auto sheet = wb->sheet("数据表");
for (int i = 0; i < 100; ++i) {
    sheet.set(i, 0, data[i]);
}
```

### 4. 合理使用类型转换

```cpp
// ❌ 低效：多次类型检查
auto cell = wb->sheet(0)->getCell(1, 1);
if (cell.isString()) {
    process(cell.getStringValue());
} else if (cell.isNumber()) {
    process(cell.getNumberValue());
}

// ✅ 高效：使用 tryGetValue
auto cell = wb->sheet(0)->getCell(1, 1);
if (auto str_val = cell.tryGetValue<std::string>()) {
    process(*str_val);
} else if (auto num_val = cell.tryGetValue<double>()) {
    process(*num_val);
}
```

## 🔧 调试和错误处理

### 1. 异常安全的代码

```cpp
try {
    auto wb = Workbook::open("data.xlsx");
    
    // 安全访问可能不存在的工作表
    try {
        auto value = wb->getValue<std::string>("不存在的表", 1, 1);
    } catch (const Exception& e) {
        std::cerr << "工作表访问错误: " << e.what() << std::endl;
    }
    
    // 或使用安全方法
    auto sheet = wb->getWorksheet("可能不存在的表");
    if (sheet) {
        auto cell = sheet->getCell(1, 1);
        auto value = cell.getValueOr<std::string>("默认值");
    }
    
} catch (const std::exception& e) {
    std::cerr << "文件操作错误: " << e.what() << std::endl;
}
```

### 2. 类型转换验证

```cpp
auto wb = Workbook::open("mixed_data.xlsx");

// 验证数据类型
for (int row = 1; row <= 100; ++row) {
    auto cell = wb->sheet(0)->getCell(row, 1);
    
    if (cell.isEmpty()) {
        std::cout << "行 " << row << " 为空" << std::endl;
        continue;
    }
    
    // 尝试转换为数字
    if (auto num_val = cell.tryGetValue<double>()) {
        std::cout << "行 " << row << " 数值: " << *num_val << std::endl;
    } else {
        std::cout << "行 " << row << " 非数值: " << 
                     cell.getValueOr<std::string>("无法读取") << std::endl;
    }
}
```

## 📚 地址格式参考

| 格式 | 说明 | 示例 |
|------|------|------|
| `A1` | 单个单元格（默认第一个工作表） | `wb->getValue<int>("A1")` |
| `Sheet1!A1` | 指定工作表的单个单元格 | `wb->setValue("Sheet1!A1", "值")` |
| `A1:C3` | 范围（默认第一个工作表） | `wb->getRange<double>("A1:C3")` |
| `Sheet1!A1:C3` | 指定工作表的范围 | `wb->setRange("Sheet1!A1:C3", values)` |
| `A:A` | 整列（A列） | `wb->getRange<std::string>("A:A")` |
| `1:1` | 整行（第1行） | `wb->getRange<int>("1:1")` |

## 🎉 迁移检查清单

从旧 API 迁移到新 API 时，请检查：

- [ ] ✅ 将 `getStringValue()` 改为 `getValue<std::string>()`
- [ ] ✅ 将 `getNumberValue()` 改为 `getValue<double>()`  
- [ ] ✅ 将 `getBooleanValue()` 改为 `getValue<bool>()`
- [ ] ✅ 使用直接访问API减少代码层级
- [ ] ✅ 将循环操作改为批量操作
- [ ] ✅ 添加异常处理和类型验证
- [ ] ✅ 使用链式调用优化代码可读性
- [ ] ✅ 考虑使用合适的打开模式优化性能

---

**快速开始**: 复制任何示例代码，替换文件名和数据即可立即使用！

*快速参考版本*：v1.0  
*最后更新*：2025-01-11