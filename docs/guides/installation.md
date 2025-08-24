# FastExcel 安装配置指南

本指南提供 FastExcel 的详细安装和配置说明。

## 📋 系统要求

### 编译器要求
- **C++17 兼容编译器**:
  - GCC 7.0+ (推荐 GCC 9+)
  - Clang 5.0+ (推荐 Clang 10+)  
  - MSVC 2017+ (推荐 MSVC 2019+)
  - AppleClang 10+ (macOS)

### 构建工具
- **CMake**: 3.15 或更高版本
- **Git**: 用于克隆仓库和子模块
- **Make** (Linux/macOS) 或 **MSBuild** (Windows)

### 平台支持
- ✅ **Windows 10/11** (x64, ARM64)
- ✅ **Ubuntu 18.04+** (x64, ARM64)
- ✅ **CentOS 8+** / **RHEL 8+**
- ✅ **macOS 10.15+** (x64, Apple Silicon)
- ✅ **Debian 10+**

## 🚀 快速安装

### 方法一：从源码编译 (推荐)

```bash
# 1. 克隆仓库
git clone --recursive https://github.com/wuxianggujun/FastExcel.git
cd FastExcel

# 2. 创建构建目录
mkdir build && cd build

# 3. 配置CMake
cmake .. -DCMAKE_BUILD_TYPE=Release

# 4. 编译
cmake --build . --parallel 4

# 5. 安装 (可选)
cmake --install . --prefix /usr/local
```

### 方法二：使用vcpkg (推荐Windows用户)

```bash
# 1. 安装vcpkg (如果尚未安装)
git clone https://github.com/Microsoft/vcpkg.git
cd vcpkg
./bootstrap-vcpkg.sh  # Linux/macOS
# 或 .\bootstrap-vcpkg.bat  # Windows

# 2. 安装FastExcel
./vcpkg install fastexcel

# 3. 集成到项目
cmake -DCMAKE_TOOLCHAIN_FILE=path/to/vcpkg/scripts/buildsystems/vcpkg.cmake ..
```

### 方法三：使用Conan

```bash
# 1. 安装Conan
pip install conan

# 2. 添加远程仓库
conan remote add fastexcel https://conan.fastexcel.org

# 3. 安装依赖
conan install fastexcel/2.0.0@ -if build

# 4. 使用CMake
cmake .. -DCMAKE_TOOLCHAIN_FILE=build/conan_toolchain.cmake
```

## 🔧 详细配置选项

### CMake 配置选项

```bash
cmake .. \
  -DCMAKE_BUILD_TYPE=Release \
  -DFASTEXCEL_BUILD_SHARED_LIBS=ON \
  -DFASTEXCEL_BUILD_EXAMPLES=ON \
  -DFASTEXCEL_BUILD_TESTS=ON \
  -DFASTEXCEL_USE_SYSTEM_LIBS=OFF \
  -DFASTEXCEL_USE_LIBDEFLATE=ON \
  -DFASTEXCEL_INSTALL=ON
```

#### 配置选项说明

| 选项 | 默认值 | 说明 |
|------|--------|------|
| `FASTEXCEL_BUILD_SHARED_LIBS` | `OFF` | 构建动态库而非静态库 |
| `FASTEXCEL_BUILD_EXAMPLES` | `ON` | 构建示例程序 |
| `FASTEXCEL_BUILD_TESTS` | `ON` | 构建单元测试 |
| `FASTEXCEL_BUILD_UNIT_TESTS` | `ON` | 构建单元测试 |
| `FASTEXCEL_BUILD_INTEGRATION_TESTS` | `OFF` | 构建集成测试 |
| `FASTEXCEL_USE_SYSTEM_LIBS` | `OFF` | 使用系统库而非捆绑库 |
| `FASTEXCEL_USE_LIBDEFLATE` | `OFF` | 启用libdeflate高性能压缩 |
| `FASTEXCEL_INSTALL` | `OFF` | 启用安装目标 |

### 构建类型选择

```bash
# 发布版本 (最优性能)
cmake .. -DCMAKE_BUILD_TYPE=Release

# 调试版本 (包含调试符号)
cmake .. -DCMAKE_BUILD_TYPE=Debug

# 带调试信息的发布版本
cmake .. -DCMAKE_BUILD_TYPE=RelWithDebInfo

# 最小大小发布版本
cmake .. -DCMAKE_BUILD_TYPE=MinSizeRel
```

## 🏗️ 平台特定安装

### Windows (MSVC)

#### 使用 Visual Studio

```bash
# 1. 克隆项目
git clone --recursive https://github.com/wuxianggujun/FastExcel.git
cd FastExcel

# 2. 生成Visual Studio项目
mkdir build && cd build
cmake .. -G "Visual Studio 16 2019" -A x64

# 3. 在Visual Studio中打开FastExcel.sln并构建
# 或使用命令行构建
cmake --build . --config Release --parallel 4
```

#### 使用 MinGW

```bash
# 1. 安装MSYS2和MinGW-w64
# 2. 在MSYS2环境中
pacman -S mingw-w64-x86_64-gcc mingw-w64-x86_64-cmake

# 3. 构建
mkdir build && cd build
cmake .. -G "MinGW Makefiles"
cmake --build . --parallel 4
```

### Linux (Ubuntu/Debian)

#### 安装依赖

```bash
# Ubuntu 20.04+
sudo apt update
sudo apt install -y \
    build-essential \
    cmake \
    git \
    pkg-config \
    libc6-dev

# 可选：安装系统库 (如果使用FASTEXCEL_USE_SYSTEM_LIBS=ON)
sudo apt install -y \
    libexpat1-dev \
    zlib1g-dev \
    libfmt-dev
```

#### 编译安装

```bash
git clone --recursive https://github.com/wuxianggujun/FastExcel.git
cd FastExcel
mkdir build && cd build
cmake .. -DCMAKE_BUILD_TYPE=Release
make -j$(nproc)
sudo make install
```

### Linux (CentOS/RHEL)

#### 安装依赖

```bash
# CentOS 8+ / RHEL 8+
sudo dnf groupinstall "Development Tools"
sudo dnf install cmake git pkgconfig

# 启用EPEL仓库获取更新的包
sudo dnf install epel-release
sudo dnf install fmt-devel expat-devel zlib-devel
```

### macOS

#### 使用 Homebrew

```bash
# 1. 安装Homebrew (如果尚未安装)
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

# 2. 安装依赖
brew install cmake git pkg-config

# 3. 编译FastExcel
git clone --recursive https://github.com/wuxianggujun/FastExcel.git
cd FastExcel
mkdir build && cd build
cmake .. -DCMAKE_BUILD_TYPE=Release
make -j$(sysctl -n hw.ncpu)
```

#### 使用 MacPorts

```bash
# 1. 安装依赖
sudo port install cmake git pkgconfig

# 2. 编译 (同上)
```

## 📦 依赖库管理

### 内置依赖 (推荐)

FastExcel 默认使用内置的第三方库，确保版本兼容性：

```
third_party/
├── fmt/                 # 字符串格式化库
├── libexpat/           # XML解析库  
├── ZlibMinizipBundle/  # ZIP压缩库
├── googletest/         # 测试框架
├── utfcpp/            # UTF-8处理
├── stb/               # 图像处理
└── libdeflate/        # 高性能压缩 (可选)
```

### 使用系统库

```bash
# 如果系统已安装兼容版本的依赖库
cmake .. -DFASTEXCEL_USE_SYSTEM_LIBS=ON

# 需要确保以下库版本兼容:
# - fmt >= 8.0
# - libexpat >= 2.2
# - zlib >= 1.2.11
# - googletest >= 1.10 (仅测试需要)
```

## 🔍 验证安装

### 运行测试

```bash
# 在build目录中
ctest --parallel 4 --output-on-failure

# 或运行特定测试
./test/fastexcel_unit_tests
./test/fastexcel_integration_tests
```

### 运行示例

```bash
# 基础示例
cd build/examples
./basic_usage_example

# 性能示例
./performance_demo

# 格式化示例
./formatting_features_demo
```

### 验证库链接

```cpp
// test_installation.cpp
#include "src/fastexcel/FastExcel.hpp"

int main() {
    std::cout << "FastExcel 版本: " << fastexcel::getVersion() << std::endl;
    
    // 创建简单工作簿测试
    auto workbook = fastexcel::core::Workbook::create("test.xlsx");
    auto worksheet = workbook->addWorksheet("测试");
    worksheet->getCell(0, 0)->setValue("Hello FastExcel!");
    
    if (workbook->save()) {
        std::cout << "安装验证成功!" << std::endl;
        return 0;
    } else {
        std::cout << "安装验证失败!" << std::endl;
        return 1;
    }
}
```

编译测试：
```bash
g++ -std=c++17 -I./src test_installation.cpp -L./build/lib -lfastexcel -o test
./test
```

## ⚙️ 性能调优

### 编译器优化选项

#### GCC 优化

```bash
cmake .. \
  -DCMAKE_BUILD_TYPE=Release \
  -DCMAKE_CXX_FLAGS="-O3 -march=native -mtune=native -flto"
```

#### Clang 优化

```bash
cmake .. \
  -DCMAKE_BUILD_TYPE=Release \
  -DCMAKE_CXX_FLAGS="-O3 -march=native -flto"
```

#### MSVC 优化

```bash
cmake .. \
  -DCMAKE_BUILD_TYPE=Release \
  -DCMAKE_CXX_FLAGS="/O2 /GL /LTCG"
```

### 链接时优化 (LTO)

```bash
# 启用链接时优化以获得更好性能
cmake .. \
  -DCMAKE_BUILD_TYPE=Release \
  -DCMAKE_INTERPROCEDURAL_OPTIMIZATION=ON
```

### 高性能配置

```bash
# 最大性能配置
cmake .. \
  -DCMAKE_BUILD_TYPE=Release \
  -DFASTEXCEL_USE_LIBDEFLATE=ON \
  -DCMAKE_INTERPROCEDURAL_OPTIMIZATION=ON \
  -DCMAKE_CXX_FLAGS="-O3 -march=native -DNDEBUG"
```

## 🐛 故障排除

### 常见编译错误

#### 1. C++17支持问题

```
错误: 'std::optional' was not declared in this scope
解决: 确保编译器支持C++17并正确设置了std标准
```

```bash
# 解决方案
cmake .. -DCMAKE_CXX_STANDARD=17 -DCMAKE_CXX_STANDARD_REQUIRED=ON
```

#### 2. 找不到依赖库

```
错误: Could not find package fmt
解决: 使用内置依赖或安装系统库
```

```bash
# 解决方案1: 使用内置依赖 (推荐)
cmake .. -DFASTEXCEL_USE_SYSTEM_LIBS=OFF

# 解决方案2: 安装系统依赖
sudo apt install libfmt-dev  # Ubuntu
brew install fmt             # macOS
```

#### 3. 内存不足

```
错误: c++: fatal error: Killed (program cc1plus)
解决: 增加swap空间或减少并行编译数
```

```bash
# 减少并行度
cmake --build . --parallel 1

# 或增加swap空间 (Linux)
sudo fallocate -l 2G /swapfile
sudo chmod 600 /swapfile
sudo mkswap /swapfile
sudo swapon /swapfile
```

### 运行时问题

#### 1. 找不到动态库

```bash
# Linux: 设置LD_LIBRARY_PATH
export LD_LIBRARY_PATH=/usr/local/lib:$LD_LIBRARY_PATH

# macOS: 设置DYLD_LIBRARY_PATH  
export DYLD_LIBRARY_PATH=/usr/local/lib:$DYLD_LIBRARY_PATH

# Windows: 将DLL路径添加到PATH环境变量
```

#### 2. 权限问题

```bash
# 确保有写入权限
chmod 755 /path/to/output/directory

# 或使用不同的输出目录
auto workbook = Workbook::create("~/Documents/output.xlsx");
```

## 📁 项目集成

### CMake项目集成

#### 方法一：add_subdirectory

```cmake
# CMakeLists.txt
cmake_minimum_required(VERSION 3.15)
project(MyProject)

# 添加FastExcel子目录
add_subdirectory(third_party/FastExcel)

# 创建你的目标
add_executable(my_app main.cpp)

# 链接FastExcel
target_link_libraries(my_app PRIVATE fastexcel)
```

#### 方法二：find_package

```cmake
# CMakeLists.txt
find_package(FastExcel REQUIRED)

add_executable(my_app main.cpp)
target_link_libraries(my_app PRIVATE FastExcel::fastexcel)
```

#### 方法三：FetchContent

```cmake
# CMakeLists.txt
include(FetchContent)

FetchContent_Declare(
    FastExcel
    GIT_REPOSITORY https://github.com/wuxianggujun/FastExcel.git
    GIT_TAG        v2.0.0
)
FetchContent_MakeAvailable(FastExcel)

target_link_libraries(my_app PRIVATE fastexcel)
```

### pkg-config 集成

```bash
# 安装后生成 fastexcel.pc 文件
pkg-config --cflags fastexcel
pkg-config --libs fastexcel

# 在Makefile中使用
CXXFLAGS += $(shell pkg-config --cflags fastexcel)
LDFLAGS += $(shell pkg-config --libs fastexcel)
```

## 📚 下一步

安装完成后，你可以：

1. 查看 [快速开始指南](../guides/quick-start.md) 学习基本用法
2. 浏览 [示例程序](../examples/) 了解各种功能
3. 阅读 [API文档](../api/) 掌握完整接口  
4. 参考 [性能指南](../architecture/performance.md) 优化应用

如果遇到问题，请查看 [FAQ](../guides/faq.md) 或在 [GitHub Issues](https://github.com/wuxianggujun/FastExcel/issues) 中报告。