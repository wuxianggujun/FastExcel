# 归档层（archive）API 与类关系

负责 XLSX 包所依赖的 ZIP 容器的读取、写入与压缩参数管理。提供 Reader/Writer/Archive 三个层级，以及压缩引擎与文件管理器等辅助。

## 关系图（简述）
- `ZipArchive` 组合 `ZipReader` 与 `ZipWriter`，对外提供读/写统一入口
- `ZipReader`/`ZipWriter` 直接操作底层压缩库（minizip 等），线程安全封装
- `core::Path` 为路径依赖；`opc::PackageEditor`、`reader::XLSXReader` 广泛使用本层

---

## enum class fastexcel::archive::ZipError
- 语义：统一错误码，`Ok` 表示成功；提供 `isSuccess/ isError` 与逻辑非运算 `!` 的便捷语义。

---

## class fastexcel::archive::ZipArchive
- 职责：在单实例中同时协调 ZIP 读取与写入能力；适合“读取部分文件后写回”的场景。
- 主要 API：
  - 生命周期：`open(create=true)`、`close()`、`isOpen()`、`isWritable()`、`isReadable()`
  - 写入：`addFile(internal_path, content|data)`、批量 `addFiles(...)`、流式 `openEntry/writeChunk/closeEntry`
  - 读取：`extractFile(internal_path, string|vector|ostream)`、`fileExists(path)`、`listFiles()`
  - 配置：`setCompressionLevel(level)`
  - 直接访问：`getReader()`、`getWriter()`

---

## class fastexcel::archive::ZipReader
- 职责：高性能 ZIP 读取器；支持条目缓存、流式读取、原始压缩数据提取（便于零拷贝转存）。
- 主要 API：
  - 生命周期：`open()`、`close()`、`isOpen()`、`getPath()`
  - 查询：`listFiles()`、`listEntriesInfo()`、`fileExists(path)`、`getEntryInfo(path, info)`
  - 读取：`extractFile(path, string|vector)`、`extractFileToStream(path, ostream)`、`streamFile(path, callback, buffer)`
  - 原始数据：`getRawCompressedData(path, raw, info)`
  - 统计：`getStats()`（条目数、压缩/未压缩总量、压缩比）

---

## class fastexcel::archive::ZipWriter
- 职责：高性能 ZIP 写入器；支持批量/流式写入、防重复、统计与压缩级别控制。
- 主要 API：
  - 生命周期：`open(create)`、`close()`、`isOpen()`、`getPath()`
  - 写入：`addFile(path, content|data|bytes)`、批量 `addFiles(vector<FileEntry>)`、流式 `openEntry/writeChunk/closeEntry()`
  - 原始数据直写：`writeRawCompressedData(path, raw, size, uncompressed_size, crc32, method)`
  - 配置：`setCompressionLevel(level)`、`getCompressionLevel()`
  - 状态：`hasEntry(path)`、`getWrittenPaths()`、`getStats()`

---

## 其他组件（概览）
- `archive::FileManager`：面向包内路径的文件添加与缓冲调度（被 Workbook/生成器使用）。
- 压缩引擎：
  - `CompressionEngine`：压缩策略抽象接口
  - `ZlibEngine` / `LibDeflateEngine`：不同后端实现（压缩/解压）
  - `MinizipParallelWriter`：并行分片写入，提升大包写入吞吐
