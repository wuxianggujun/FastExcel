# FastExcel 项目 AI 编码规范

## 1. 概述

本文档为 FastExcel 项目的 AI 编码规范，确保所有 AI 生成的代码符合项目的技术要求和代码风格。FastExcel 是一个基于 C++17 的高性能 Excel 文件处理库，使用 minizip-ng、zlib-ng、libexpat、spdlog 等依赖库实现。

## 2. C++ 语言规范

### 2.1 C++ 标准

- **必须使用 C++17 标准**
- 项目在 [`CMakeLists.txt`](CMakeLists.txt:5) 中明确设置：`set(CMAKE_CXX_STANDARD 17)`
- 所有代码必须兼容 C++17，可以使用 C++17 的特性（如 `std::variant`、`std::optional`、结构化绑定等）

### 2.2 头文件规范

- **头文件扩展名必须使用 `.hpp`**
- **源文件扩展名必须使用 `.cpp`**
- **禁止使用 `.h` 作为头文件扩展名**

#### 头文件结构示例：

```cpp
#pragma once

// 标准库头文件
#include <string>
#include <vector>
#include <memory>

// 项目内部头文件
#include "fastexcel/core/Cell.hpp"
#include "fastexcel/core/Format.hpp"

// 第三方库头文件
#include <spdlog/spdlog.h>

namespace fastexcel {
namespace core {

// 类定义
class Example {
private:
    // 成员变量
    std::string name_;
    
public:
    // 构造函数和析构函数
    explicit Example(const std::string& name);
    ~Example() = default;
    
    // 成员函数
    void setName(const std::string& name);
    std::string getName() const;
};

}} // namespace fastexcel::core
```

## 3. 代码风格规范

### 3.1 命名约定

#### 文件命名
- 头文件：`PascalCase.hpp`（如 `Cell.hpp`、`Workbook.hpp`）
- 源文件：`PascalCase.cpp`（如 `Cell.cpp`、`Workbook.cpp`）

#### 类命名
- 使用 `PascalCase`
- 示例：`class Workbook`、`class XMLWriter`

#### 函数命名
- 成员函数：`camelCase`
- 示例：`getValue()`、`setName()`

#### 变量命名
- 成员变量：`trailing_underscore_`
- 局部变量：`camelCase`
- 常量：`UPPER_SNAKE_CASE`
- 示例：
  ```cpp
  class Example {
  private:
      std::string name_;        // 成员变量
      int max_count_;           // 成员变量
      
  public:
      void process() {
          int itemCount = 0;    // 局部变量
          const int MAX_ITEMS = 100;  // 常量
      }
  };
  ```

#### 命名空间
- 使用 `lowercase`
- 示例：`namespace fastexcel`、`namespace core`

### 3.2 代码格式

#### 缩进和空白
- 使用 4 个空格缩进，**禁止使用制表符**
- 大括号使用 K&R 风格：
  ```cpp
  class Example {
  public:
      void method() {
          if (condition) {
              // 代码
          }
      }
  };
  ```

#### 行长度
- 建议每行不超过 120 个字符
- 长行应在逻辑位置换行

#### 空行使用
- 函数之间：1-2 个空行
- 类内部：逻辑块之间 1 个空行
- 不要连续使用超过 2 个空行

### 3.3 注释规范

#### 文件头注释
```cpp
#pragma once

// FastExcel 库 - 高性能Excel文件处理库
// 组件：核心单元格处理模块
//
// 本文件实现了 Excel 单元格的基本功能，包括数据存储、类型转换和格式应用。
```

#### 类注释
```cpp
/**
 * @brief Excel 单元格类
 * 
 * Cell 类表示 Excel 工作表中的单个单元格，支持多种数据类型
 * 包括字符串、数字、布尔值、日期和公式。
 */
class Cell {
    // ...
};
```

#### 函数注释
```cpp
/**
 * @brief 设置单元格的字符串值
 * @param value 要设置的字符串值
 * 
 * 将单元格的类型设置为 String，并存储指定的字符串值。
 * 如果之前设置了公式，公式将被清除。
 */
void setValue(const std::string& value);
```

## 4. 架构规范

### 4.1 项目结构

```
FastExcel/
├── src/fastexcel/           # 源代码目录
│   ├── core/               # 核心功能模块
│   │   ├── Cell.hpp/cpp
│   │   ├── Format.hpp/cpp
│   │   ├── Workbook.hpp/cpp
│   │   └── Worksheet.hpp/cpp
│   ├── xml/                # XML 处理模块
│   │   ├── XMLWriter.hpp/cpp
│   │   ├── ContentTypes.hpp/cpp
│   │   ├── Relationships.hpp/cpp
│   │   └── SharedStrings.hpp/cpp
│   ├── archive/            # 压缩归档模块
│   │   ├── ZipArchive.hpp/cpp
│   │   └── FileManager.hpp/cpp
│   ├── utils/              # 工具模块
│   │   └── Logger.hpp/cpp
│   └── FastExcel.hpp      # 主头文件
├── include/fastexcel/     # 公共头文件目录（目前为空）
├── examples/              # 示例程序
├── docs/                  # 文档目录
├── third_party/           # 第三方依赖
└── CMakeLists.txt         # 构建配置
```

### 4.2 模块依赖关系

- **核心层 (core)**：不依赖其他业务模块，只依赖标准库和工具层
- **XML 层 (xml)**：依赖工具层，不依赖核心层和压缩层
- **压缩层 (archive)**：依赖工具层，不依赖核心层和 XML 层
- **工具层 (utils)**：独立模块，只依赖第三方库

### 4.3 命名空间使用

- 所有代码必须位于 `fastexcel` 命名空间下
- 子模块使用子命名空间：
  ```cpp
  namespace fastexcel {
  namespace core {
      class Cell { /* ... */ };
  }
  
  namespace xml {
      class XMLWriter { /* ... */ };
  }
  
  namespace archive {
      class ZipArchive { /* ... */ };
  }
  }
  ```

## 5. 技术规范

### 5.1 内存管理

- 优先使用智能指针 (`std::unique_ptr`, `std::shared_ptr`)
- 避免原始指针的显式 `new`/`delete`
- 使用 RAII 模式管理资源

#### 示例：
```cpp
// 推荐：使用智能指针
std::shared_ptr<Worksheet> worksheet = std::make_shared<Worksheet>("Sheet1");
std::unique_ptr<FileManager> file_manager = std::make_unique<FileManager>();

// 不推荐：原始指针
Worksheet* worksheet = new Worksheet("Sheet1");
// ... 需要手动 delete
delete worksheet;
```

### 5.2 错误处理

- 使用异常处理错误情况
- 自定义异常类应继承自 `std::exception`
- 日志记录重要操作和错误

#### 示例：
```cpp
try {
    auto workbook = std::make_shared<fastexcel::core::Workbook>("example.xlsx");
    if (!workbook->open()) {
        LOG_ERROR("Failed to open workbook");
        return false;
    }
    // ... 其他操作
} catch (const std::exception& e) {
    LOG_ERROR("Exception occurred: {}", e.what());
    return false;
}
```

### 5.3 日志规范

- 使用 spdlog 库进行日志记录
- 使用项目定义的日志宏：
  ```cpp
  LOG_TRACE("详细追踪信息");
  LOG_DEBUG("调试信息");
  LOG_INFO("一般信息");
  LOG_WARN("警告信息");
  LOG_ERROR("错误信息");
  LOG_CRITICAL("严重错误信息");
  ```

### 5.4 常量定义

- 使用 `constexpr` 定义编译时常量
- 使用 `enum class` 定义枚举类型
- 避免使用宏定义常量（特殊情况除外）

#### 示例：
```cpp
// 推荐：使用 constexpr
constexpr int MAX_WORKSHEETS = 255;
constexpr double PI = 3.14159265359;

// 推荐：使用 enum class
enum class CellType {
    Empty,
    String,
    Number,
    Boolean,
    Date,
    Formula,
    Error
};

// 不推荐：使用宏定义
#define MAX_WORKSHEETS 255  // 避免这种方式
```

## 6. 性能规范

### 6.1 字符串处理

- 优先使用 `std::string_view` 进行只读字符串操作
- 避免不必要的字符串拷贝
- 使用字符串移动语义 (`std::move`)

#### 示例：
```cpp
// 推荐：使用 string_view
void processName(std::string_view name) {
    // 处理名称，不涉及拷贝
}

// 不推荐：值传递
void processName(std::string name) {
    // 会产生字符串拷贝
}
```

### 6.2 容器使用

- 选择合适的容器类型：
  - `std::vector`：连续内存，随机访问
  - `std::map`：有序键值对
  - `std::unordered_map`：无序键值对，查找更快
- 预分配容器容量以减少重新分配

#### 示例：
```cpp
// 推荐：预分配容量
std::vector<std::string> names;
names.reserve(1000);  // 预分配容量

// 不推荐：频繁重新分配
std::vector<std::string> names;
for (int i = 0; i < 1000; ++i) {
    names.push_back(getName(i));  // 可能导致多次重新分配
}
```

### 6.3 算法复杂度

- 关注算法的时间复杂度和空间复杂度
- 优先使用标准库算法 (`<algorithm>`)
- 避免嵌套循环，考虑使用更高效的算法

## 7. 测试规范

### 7.1 测试框架

- **必须使用 Gtest 测试框架**
- 禁止使用 Catch2 或其他测试框架
- 所有测试文件必须放在 `test/` 目录下：
  - `test/unit/` - 单元测试
  - `test/integration/` - 集成测试
- 禁止 Gtest 框架自身的测试构建选项：
  ```cmake
  set(gtest_build_tests OFF CACHE BOOL "" FORCE)
  set(gtest_build_samples OFF CACHE BOOL "" FORCE)
  set(gtest_disable_pthreads ON CACHE BOOL "" FORCE)
  ```
- 使用以下选项控制测试构建：
  ```cmake
  option(ENABLE_TESTS "Enable building of all tests" OFF)
  option(ENABLE_UNIT_TESTS "Enable building of unit tests" OFF)
  option(ENABLE_INTEGRATION_TESTS "Enable building of integration tests" OFF)
  ```
  - `ENABLE_TESTS` 是总开关，开启后会自动启用所有测试类型
  - `ENABLE_UNIT_TESTS` 和 `ENABLE_INTEGRATION_TESTS` 可以单独控制特定类型的测试
  - 当 `ENABLE_TESTS=ON` 时，会自动设置 `ENABLE_UNIT_TESTS=ON` 和 `ENABLE_INTEGRATION_TESTS=ON`

### 7.2 单元测试

- 每个类都应该有对应的单元测试
- 测试文件命名：`ClassName_test.cpp`
- 测试类命名：`ClassNameTest`
- 测试函数命名：`TEST_F(ClassNameTest, 功能描述)`
- 使用 Gtest 的断言宏：`EXPECT_EQ`, `ASSERT_TRUE`, 等

#### 示例：
```cpp
// Cell_test.cpp
#include "fastexcel/core/Cell.hpp"
#include <gtest/gtest.h>

namespace fastexcel {
namespace core {

class CellTest : public ::testing::Test {
protected:
    void SetUp() override {
        cell = std::make_unique<Cell>();
    }

    void TearDown() override {
        cell.reset();
    }

    std::unique_ptr<Cell> cell;
};

TEST_F(CellTest, StringValue) {
    cell->setValue("Hello");
    EXPECT_EQ(cell->getType(), CellType::String);
    EXPECT_EQ(cell->getStringValue(), "Hello");
}

TEST_F(CellTest, NumberValue) {
    cell->setValue(42.0);
    EXPECT_EQ(cell->getType(), CellType::Number);
    EXPECT_DOUBLE_EQ(cell->getNumberValue(), 42.0);
}

}} // namespace fastexcel::core
```

### 7.3 集成测试

- 测试模块间的交互
- 测试文件放在 `test/integration/` 目录
- 测试文件命名：`ModuleIntegration_test.cpp`
- 测试完整的业务流程
- 使用 Gtest 的测试夹具（Test Fixtures）管理测试资源

#### 示例：
```cpp
// WorkbookIntegration_test.cpp
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/Worksheet.hpp"
#include <gtest/gtest.h>

namespace fastexcel {
namespace core {

class WorkbookIntegrationTest : public ::testing::Test {
protected:
    void SetUp() override {
        workbook = std::make_unique<Workbook>("test_workbook.xlsx");
    }

    void TearDown() override {
        workbook.reset();
    }

    std::unique_ptr<Workbook> workbook;
};

TEST_F(WorkbookIntegrationTest, CreateWorkbookAndAddWorksheet) {
    auto worksheet = workbook->addWorksheet("Sheet1");
    EXPECT_NE(worksheet, nullptr);
    EXPECT_EQ(worksheet->getName(), "Sheet1");
}

}} // namespace fastexcel::core
```

## 8. 文档规范

### 8.1 代码文档

- 所有公共 API 必须有完整的文档注释
- 使用 Doxygen 风格的注释
- 复杂算法必须有详细的实现说明

### 8.2 用户文档

- 在 `docs/` 目录下提供用户指南
- 提供 API 参考文档
- 包含示例代码和最佳实践

## 9. 构建规范

### 9.1 CMake 规范

- 使用现代 CMake 语法 (3.10+)
- 遵循项目的 [`CMakeLists.txt`](CMakeLists.txt:1) 结构
- 明确指定依赖关系

#### 示例：
```cmake
# 设置 C++ 标准
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# 添加库
add_library(fastexcel STATIC ${SOURCES})

# 设置包含目录
target_include_directories(fastexcel PUBLIC
    ${CMAKE_CURRENT_SOURCE_DIR}/src
)

# 链接依赖库
target_link_libraries(fastexcel PUBLIC
    spdlog::spdlog
    expat::expat
    zlib
    MINIZIP::minizip
)
```

### 9.2 编译器警告

- 启用严格的编译器警告
- 将警告视为错误（在开发阶段）
- 使用统一的编译器设置

## 10. 版本控制规范

### 10.1 Git 工作流

- 使用功能分支开发
- 提交信息清晰明了
- 定期同步主分支

### 10.2 代码审查

- 所有代码变更必须经过审查
- 自动化检查必须通过
- 人工审查关注代码质量和设计

## 11. 安全规范

### 11.1 输入验证

- 所有外部输入必须验证
- 防止缓冲区溢出
- 处理异常输入情况

### 11.2 资源管理

- 确保所有资源正确释放
- 避免资源泄漏
- 使用 RAII 模式

## 12. 兼容性规范

### 12.1 平台兼容性

- 支持 Windows、Linux、macOS
- 使用平台抽象层
- 避免平台特定代码

### 12.2 编译器兼容性

- 支持 MSVC、GCC、Clang
- 使用标准 C++ 特性
- 避免编译器特定扩展

## 13. 终端命令规范

### 13.1 PowerShell 命令规则

- **必须使用 PowerShell 命令**，禁止使用 cmd 命令
- 所有终端命令示例必须使用 PowerShell 语法
- 使用 PowerShell 的标准 cmdlet 和参数格式

#### 基本命令规范

##### 目录操作
```powershell
# 创建目录
New-Item -ItemType Directory -Path "test\unit"

# 删除目录
Remove-Item -Path "build" -Recurse -Force

# 复制目录
Copy-Item -Path "src" -Destination "backup" -Recurse

# 移动目录
Move-Item -Path "old_folder" -Destination "new_folder"
```

##### 文件操作
```powershell
# 创建文件
New-Item -ItemType File -Path "test\unit\Cell_test.cpp"

# 删除文件
Remove-Item -Path "*.tmp" -Force

# 复制文件
Copy-Item -Path "config.json" -Destination "backup\"

# 移动文件
Move-Item -Path "old_name.cpp" -Destination "new_name.cpp"
```

##### 构建命令
```powershell
# 创建构建目录
New-Item -ItemType Directory -Path "build"

# 进入构建目录
Set-Location -Path "build"

# 配置 CMake
cmake .. -G "Visual Studio 17 2022" -A x64

# 构建项目
cmake --build . --config Release

# 运行测试
ctest -C Release -V

# 返回根目录
Set-Location -Path ".."
```

##### Git 操作
```powershell
# 克隆仓库
git clone https://github.com/username/repository.git

# 添加文件
git add .

# 提交更改
git commit -m "Commit message"

# 推送到远程仓库
git push origin main

# 拉取最新更改
git pull origin main
```

##### 环境变量操作
```powershell
# 设置环境变量
$env:PATH += ";C:\path\to\directory"

# 获取环境变量
$env:PATH

# 永久设置环境变量
[Environment]::SetEnvironmentVariable("PATH", $env:PATH, "User")
```

##### 进程管理
```powershell
# 查看进程
Get-Process

# 停止进程
Stop-Process -Name "notepad" -Force

# 启动进程
Start-Process -FilePath "notepad.exe"
```

### 13.2 命令注释规范

- 所有复杂的 PowerShell 命令必须包含注释
- 注释使用 `#` 符号
- 注释应该解释命令的目的和重要参数

#### 示例：
```powershell
# 创建测试目录结构
New-Item -ItemType Directory -Path "test\unit" -Force
New-Item -ItemType Directory -Path "test\integration" -Force

# 配置 CMake 项目，禁用不必要的测试
cmake .. -G "Visual Studio 17 2022" -A x64 `
    -DSPDLOG_BUILD_TESTS=OFF `
    -DEXPAT_BUILD_TESTS=OFF `
    -Dgtest_build_tests=OFF

# 构建项目并运行测试
cmake --build . --config Release
ctest -C Release -V --output-on-failure
```

## 14. 总结

本规范涵盖了 FastExcel 项目的各个方面，确保 AI 生成的代码符合项目的技术要求和代码风格。所有 AI 参与代码生成时，必须严格遵守本规范。

### 关键要点回顾：

1. **C++17 标准**：必须使用 C++17，禁止使用其他标准
2. **头文件扩展名**：必须使用 `.hpp`，禁止使用 `.h`
3. **源文件扩展名**：必须使用 `.cpp`
4. **命名约定**：遵循项目既定的命名规范
5. **代码风格**：使用 4 空格缩进，K&R 大括号风格
6. **架构设计**：遵循模块化设计，明确依赖关系
7. **内存管理**：优先使用智能指针，避免原始指针
8. **错误处理**：使用异常处理和日志记录
9. **性能考虑**：关注算法复杂度和资源使用
10. **测试覆盖**：确保单元测试和集成测试的完整性
11. **测试框架**：必须使用 Gtest，禁止使用 Catch2
12. **终端命令**：必须使用 PowerShell，禁止使用 cmd

遵循这些规范将确保 FastExcel 项目的高质量、可维护性和一致性。