# FastExcel 构建指南

## 构建选项

FastExcel 提供了灵活的构建配置选项，支持不同的使用场景。

### 主要构建选项

| 选项 | 默认值 | 说明 |
|------|--------|------|
| `FASTEXCEL_BUILD_SHARED_LIBS` | `ON` | 构建共享库（DLL/SO）而非静态库 |
| `FASTEXCEL_USE_SYSTEM_LIBS` | `OFF` | 使用系统已安装的依赖库 |
| `FASTEXCEL_BUILD_EXAMPLES` | `OFF` | 构建示例程序 |
| `FASTEXCEL_BUILD_TESTS` | `ON` | 构建测试程序 |
| `FASTEXCEL_BUILD_UNIT_TESTS` | `ON` | 构建单元测试 |
| `FASTEXCEL_BUILD_INTEGRATION_TESTS` | `OFF` | 构建集成测试 |

## 构建示例

### 1. 默认构建（共享库）
```bash
cmake -B build -S .
cmake --build build
```

### 2. 构建静态库
```bash
cmake -B build -S . -DFASTEXCEL_BUILD_SHARED_LIBS=OFF
cmake --build build
```

### 3. 使用系统库构建
```bash
cmake -B build -S . -DFASTEXCEL_USE_SYSTEM_LIBS=ON
cmake --build build
```

### 4. 构建示例和测试
```bash
cmake -B build -S . \
    -DFASTEXCEL_BUILD_EXAMPLES=ON \
    -DFASTEXCEL_BUILD_TESTS=ON
cmake --build build
```

### 5. 发布版本构建
```bash
cmake -B build -S . \
    -DCMAKE_BUILD_TYPE=Release \
    -DFASTEXCEL_BUILD_SHARED_LIBS=ON \
    -DFASTEXCEL_BUILD_TESTS=OFF
cmake --build build
```

## 库大小优化

### 共享库 vs 静态库

- **共享库** (`FASTEXCEL_BUILD_SHARED_LIBS=ON`)
  - ✅ 库文件较小（~2-5MB）
  - ✅ 多个程序可共享同一库
  - ✅ 更新库时无需重新编译程序
  - ❌ 需要分发 DLL/SO 文件

- **静态库** (`FASTEXCEL_BUILD_SHARED_LIBS=OFF`)
  - ✅ 单一可执行文件，无外部依赖
  - ✅ 部署简单
  - ❌ 可执行文件较大（~10-15MB）
  - ❌ 每个程序都包含库的完整副本

### 使用系统库

启用 `FASTEXCEL_USE_SYSTEM_LIBS=ON` 可以：
- 减少编译时间
- 减少最终库的大小
- 利用系统优化的库版本

**前提条件**：系统需要安装以下库
- spdlog
- libexpat
- zlib
- GoogleTest（仅测试时需要）

## 依赖库管理

### 第三方库配置

所有第三方库的配置都集中在 `third_party/CMakeLists.txt` 中，包括：

- **spdlog**: 高性能日志库
- **libexpat**: XML 解析库
- **zlib-ng**: 高性能压缩库
- **minizip-ng**: ZIP 文件处理库
- **GoogleTest**: 单元测试框架

### 自定义第三方库版本

如果需要使用特定版本的第三方库，可以：

1. 更新 `third_party/` 目录中的子模块
2. 修改 `third_party/CMakeLists.txt` 中的配置
3. 重新构建项目

## 安装

### 系统安装
```bash
cmake -B build -S . -DCMAKE_INSTALL_PREFIX=/usr/local
cmake --build build
sudo cmake --install build
```

### 用户安装
```bash
cmake -B build -S . -DCMAKE_INSTALL_PREFIX=$HOME/.local
cmake --build build
cmake --install build
```

## 在其他项目中使用

### 使用 find_package
```cmake
find_package(FastExcel REQUIRED)
target_link_libraries(your_target PRIVATE FastExcel::fastexcel)
```

### 使用 add_subdirectory
```cmake
add_subdirectory(path/to/FastExcel)
target_link_libraries(your_target PRIVATE fastexcel)
```

## 代码质量和安全性 🆕

### 内存安全特性

FastExcel 采用现代C++最佳实践，确保内存安全：

- **智能指针架构**: 全面使用 `std::unique_ptr` 替代 raw pointers
- **RAII 资源管理**: 自动内存管理，消除内存泄漏风险
- **异常安全**: 强异常安全保证，确保资源正确释放
- **线程安全**: 智能指针提供更好的多线程安全性

### 性能优化

- **XML流式生成**: 使用 `XMLStreamWriter` 替代字符串拼接，减少内存分配
- **统一文本转义**: `XMLUtils::escapeXML()` 提供优化的XML转义处理
- **零拷贝设计**: 尽可能避免不必要的数据复制

### 构建配置建议

对于不同使用场景的推荐构建配置：

#### 开发调试版本
```bash
cmake -B build -S . \
    -DCMAKE_BUILD_TYPE=Debug \
    -DFASTEXCEL_BUILD_EXAMPLES=ON \
    -DFASTEXCEL_BUILD_TESTS=ON
```

#### 性能测试版本  
```bash
cmake -B build -S . \
    -DCMAKE_BUILD_TYPE=RelWithDebInfo \
    -DFASTEXCEL_BUILD_EXAMPLES=ON
```

#### 生产发布版本
```bash
cmake -B build -S . \
    -DCMAKE_BUILD_TYPE=Release \
    -DFASTEXCEL_BUILD_SHARED_LIBS=ON \
    -DFASTEXCEL_BUILD_TESTS=OFF
```

## 故障排除

### 常见问题

1. **找不到依赖库**
   ```bash
   # 确保子模块已初始化
   git submodule update --init --recursive
   ```

2. **编译错误**
   ```bash
   # 清理构建目录
   rm -rf build
   cmake -B build -S .
   ```

3. **链接错误**
   ```bash
   # 检查是否启用了正确的构建选项
   cmake -B build -S . -DFASTEXCEL_BUILD_SHARED_LIBS=OFF
   ```

### 调试构建

启用详细输出：
```bash
cmake --build build --verbose
```

查看配置信息：
```bash
cmake -B build -S . -DCMAKE_VERBOSE_MAKEFILE=ON