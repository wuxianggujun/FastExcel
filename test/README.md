# FastExcel 测试套件

本目录包含 FastExcel 库的完整测试套件，包括单元测试和集成测试。

## 测试结构

```
test/
├── CMakeLists.txt          # 测试构建配置
├── README.md              # 本文档
├── unit/                  # 单元测试
│   ├── test_main.cpp      # 测试主程序
│   ├── test_cell.cpp      # Cell 类测试
│   ├── test_format.cpp    # Format 类测试
│   ├── test_workbook.cpp  # Workbook 类测试
│   ├── test_worksheet.cpp # Worksheet 类测试
│   └── test_xml_writer.cpp # XMLStreamWriter 类测试
└── integration/           # 集成测试
    └── test_integration.cpp # 完整功能集成测试
```

## 测试覆盖范围

### 单元测试

#### Cell 类测试 (`test_cell.cpp`)
- ✅ 默认构造函数
- ✅ 字符串值设置和获取
- ✅ 数字值设置和获取
- ✅ 布尔值设置和获取
- ✅ 公式设置和获取
- ✅ 超链接功能
- ✅ 格式设置和获取
- ✅ 清空单元格
- ✅ 拷贝构造函数和赋值操作符
- ✅ 移动语义
- ✅ 类型转换边界情况
- ✅ 特殊数值处理

#### Format 类测试 (`test_format.cpp`)
- ✅ 默认构造函数
- ✅ 字体设置（名称、大小、颜色、样式）
- ✅ 下划线和删除线设置
- ✅ 上标和下标设置
- ✅ 对齐设置（水平、垂直、换行、旋转、缩进）
- ✅ 边框设置（样式、颜色、对角线）
- ✅ 填充设置（背景色、前景色、图案）
- ✅ 数字格式设置
- ✅ 保护设置（锁定、隐藏）
- ✅ XML 生成
- ✅ 拷贝构造和赋值操作
- ✅ 边界值和异常情况
- ✅ 复杂格式组合

#### Workbook 类测试 (`test_workbook.cpp`)
- ✅ 工作簿创建和生命周期管理
- ✅ 工作表管理（添加、获取、计数）
- ✅ 格式管理（创建、获取）
- ✅ 文档属性设置
- ✅ 自定义属性设置
- ✅ 定义名称管理
- ✅ VBA 项目支持
- ✅ 常量内存模式
- ✅ XML 生成
- ✅ 文件保存
- ✅ 错误处理
- ✅ 内存管理
- ✅ 线程安全（基本测试）

#### Worksheet 类测试 (`test_worksheet.cpp`)
- ✅ 工作表创建
- ✅ 数据写入（字符串、数字、布尔值、公式、日期、超链接）
- ✅ 批量数据写入
- ✅ 行列操作（宽度、高度、格式、隐藏）
- ✅ 合并单元格
- ✅ 自动筛选
- ✅ 冻结窗格和分割窗格
- ✅ 打印设置
- ✅ 工作表保护
- ✅ 视图设置
- ✅ 清空操作
- ✅ 行列插入删除
- ✅ XML 生成
- ✅ 参数验证
- ✅ 大量数据处理
- ✅ 性能测试

#### XMLStreamWriter 类测试 (`test_xml_writer.cpp`)
- ✅ 基本 XML 文档生成
- ✅ 元素嵌套
- ✅ 属性写入（字符串、数字）
- ✅ 空元素和自闭合元素
- ✅ 字符转义（&, <, >, ", ', \n）
- ✅ 原始数据写入
- ✅ 缓冲模式和文件模式
- ✅ 清空操作
- ✅ 属性批处理
- ✅ 大量数据处理
- ✅ 错误处理
- ✅ 性能测试
- ✅ 内存使用测试
- ✅ 线程安全（基本测试）

### 集成测试

#### 完整工作流程测试 (`test_integration.cpp`)
- ✅ 完整的 Excel 文件生成流程
- ✅ 多工作表场景
- ✅ 大数据量处理
- ✅ 复杂格式化
- ✅ 超链接功能
- ✅ 打印设置
- ✅ 工作表保护
- ✅ 批量数据写入
- ✅ 错误恢复
- ✅ 内存管理
- ✅ 并发访问（基本测试）
- ✅ 性能基准测试

## 构建和运行测试

### 前提条件

1. **GoogleTest**: 确保系统中安装了 GoogleTest 框架
2. **CMake**: 版本 3.15 或更高
3. **C++17 编译器**: GCC 7+, Clang 5+, 或 MSVC 2017+

### 构建测试

```bash
# 在项目根目录下
mkdir build
cd build
cmake .. -DBUILD_TESTING=ON
make fastexcel_tests

# 或者在 Windows 上
cmake .. -DBUILD_TESTING=ON
cmake --build . --target fastexcel_tests
```

### 运行测试

```bash
# 运行所有测试
./fastexcel_tests

# 运行特定测试套件
./fastexcel_tests --gtest_filter="CellTest.*"
./fastexcel_tests --gtest_filter="FormatTest.*"
./fastexcel_tests --gtest_filter="WorkbookTest.*"
./fastexcel_tests --gtest_filter="WorksheetTest.*"
./fastexcel_tests --gtest_filter="XMLStreamWriterTest.*"
./fastexcel_tests --gtest_filter="IntegrationTest.*"

# 运行特定测试
./fastexcel_tests --gtest_filter="CellTest.StringValue"

# 显示详细输出
./fastexcel_tests --gtest_verbose

# 生成 XML 报告
./fastexcel_tests --gtest_output=xml:test_results.xml
```

### 使用 CMake 运行测试

```bash
# 使用 CTest 运行所有测试
ctest

# 显示详细输出
ctest --verbose

# 运行特定测试
ctest -R "Cell"

# 并行运行测试
ctest -j4
```

## 测试指标

### 代码覆盖率目标
- **单元测试覆盖率**: > 90%
- **集成测试覆盖率**: > 80%
- **总体覆盖率**: > 85%

### 性能基准
- **小数据集** (100x10): < 100ms
- **中等数据集** (1000x10): < 1s
- **大数据集** (5000x5): < 5s
- **内存使用**: 常量内存模式下 < 100MB

### 质量指标
- **所有单元测试**: 必须通过
- **所有集成测试**: 必须通过
- **内存泄漏**: 零容忍
- **线程安全**: 基本并发测试通过

## 测试数据

测试过程中会生成临时文件，这些文件会在测试完成后自动清理：

- `test_workbook.xlsx` - Workbook 测试文件
- `integration_test.xlsx` - 集成测试文件
- `test_output.xml` - XML 写入器测试文件
- `memory_test_*.xlsx` - 内存管理测试文件
- `concurrent_test_*.xlsx` - 并发测试文件

## 故障排除

### 常见问题

1. **GoogleTest 未找到**
   ```
   解决方案: 安装 GoogleTest 或设置 GTEST_ROOT 环境变量
   ```

2. **编译错误**
   ```
   检查 C++17 编译器支持
   确保所有依赖库已正确安装
   ```

3. **测试失败**
   ```
   检查文件权限
   确保有足够的磁盘空间
   查看详细错误信息
   ```

4. **性能测试超时**
   ```
   可能是系统性能问题
   检查是否有其他程序占用资源
   调整性能基准阈值
   ```

### 调试测试

```bash
# 使用调试器运行测试
gdb ./fastexcel_tests
(gdb) run --gtest_filter="FailingTest.*"

# 在 Windows 上使用 Visual Studio 调试器
devenv fastexcel_tests.exe
```

### 内存检查

```bash
# 使用 Valgrind 检查内存泄漏（Linux）
valgrind --leak-check=full ./fastexcel_tests

# 使用 AddressSanitizer
cmake .. -DCMAKE_CXX_FLAGS="-fsanitize=address"
make fastexcel_tests
./fastexcel_tests
```

## 贡献测试

### 添加新测试

1. **单元测试**: 在相应的 `test_*.cpp` 文件中添加新的 `TEST_F` 函数
2. **集成测试**: 在 `test_integration.cpp` 中添加新的测试场景
3. **测试命名**: 使用描述性的测试名称，如 `TEST_F(ClassTest, SpecificFeature)`

### 测试最佳实践

1. **独立性**: 每个测试应该独立运行，不依赖其他测试
2. **可重复性**: 测试结果应该是确定性的和可重复的
3. **清理**: 测试后清理所有临时文件和资源
4. **断言**: 使用适当的 GoogleTest 断言宏
5. **文档**: 为复杂测试添加注释说明

### 测试审查清单

- [ ] 测试覆盖了所有公共接口
- [ ] 测试了边界条件和错误情况
- [ ] 测试名称清晰描述了测试内容
- [ ] 测试运行时间合理（< 1秒/测试）
- [ ] 测试后正确清理资源
- [ ] 测试通过且稳定

---

**FastExcel 测试套件** - 确保代码质量和功能正确性的重要保障！

*最后更新: 2024-01-01*