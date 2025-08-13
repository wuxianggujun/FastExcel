# FastExcel API 文档索引（v2.x）

本目录提供 FastExcel 各模块的类 API、职责说明与类之间的关系图谱。建议先阅读本页的模块概览，再进入各模块的详细文档。

- 模块概览：
  - 核心层（core）：工作簿/工作表/单元格、样式系统、范围格式化、共享字符串等。见 `core.md`。
  - 归档层（archive）：ZIP 读写与压缩编解码。见 `archive.md`。
  - OPC 包编辑（opc）：面向 XLSX 包的编辑器、包管理、部件图。见 `opc.md`。
  - XML 生成（xml）：统一 XML 生成编排器与部件生成器、流式读写器。见 `xml.md`。
  - 读取器（reader）：XLSXReader 解析工作簿、工作表、样式、共享字符串、主题等。见 `reader.md`。
  - 主题（theme）：主题颜色/字体方案与主题 XML。见 `theme.md`。
  - 变更跟踪（tracking）：包部件脏标记与变更统计。见 `tracking.md`。
  - 工具（utils）：日志、路径、地址解析、时间/通用工具、XML 工具。见 `utils.md`。

关系概览（高层）：
- `core::Workbook` 持有多个 `core::Worksheet`，聚合 `core::FormatRepository`、`core::SharedStringTable`、可选 `theme::Theme`。
- `xml::UnifiedXMLGenerator` 读取 `Workbook/Worksheet/FormatRepository/SST` 生成 XML，输出交由 `core::IFileWriter`。
- `opc::PackageEditor` 作为编排者：调用 `xml::UnifiedXMLGenerator` 写入，使用 `archive::ZipReader/ZipWriter` 管理包，借助 `tracking::IChangeTracker` 管理变更。
- `reader::XLSXReader` 借助 `archive::ZipArchive` 解析包中文件，填充 `core::Workbook/Worksheet`、样式与主题。

说明：仓库采用 `docs` 作为文档根目录，因此本 API 文档位于 `docs/api/`（而非 `doc/api/`）。

