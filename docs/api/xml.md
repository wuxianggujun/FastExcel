# XML 生成层（xml）API 与类关系

统一负责将 `core::Workbook/Worksheet` 与样式、共享字符串等信息序列化为 OOXML 各部件的 XML。

## 关系图（简述）
- `UnifiedXMLGenerator` 为编排器：
  - 输入：`GenerationContext`（Workbook/Worksheet、FormatRepository、SharedStringTable 等）
  - 输出：通过 `core::IFileWriter` 写入目标（可流式/批量）
  - 组合：多个 `IXMLPartGenerator` 子生成器（workbook、worksheets、rels、content-types、styles、sharedStrings、theme 等）
- 依赖：使用 `XMLStreamWriter` 进行流式 XML 拼写；被 `opc::PackageEditor` 调用

---

## class fastexcel::xml::UnifiedXMLGenerator
- 职责：集中管理所有部件的 XML 生成，消除重复逻辑，支持全量/增量选择性生成。
- 主要 API：
  - 构造：`UnifiedXMLGenerator(const GenerationContext&)`
  - 生成：`generateAll(IFileWriter&)`、`generateAll(IFileWriter&, const DirtyManager*)`、`generateParts(IFileWriter&, parts)`（可带 `DirtyManager`）
  - 工厂：`fromWorkbook(const core::Workbook*)`、`fromWorksheet(const core::Worksheet*)`
- 关系：内部注册 `IXMLPartGenerator` 列表并调度；对外仅暴露统一入口。

---

## class fastexcel::xml::XMLGeneratorFactory
- 职责：简化 `UnifiedXMLGenerator` 的创建；提供轻量/默认生成器的工厂方法。
- 主要 API：
  - `createGenerator(const core::Workbook*)`
  - `createLightweightGenerator()`

---

## 流式读写工具
- `XMLStreamWriter`：高效写 XML；被各部件生成器使用。
- `XMLStreamReader`：解析 XML（在 reader 侧也会使用）。

---

## 其他 XML 部件（概览）
- `ContentTypes`：`[Content_Types].xml` 生成与管理
- `Relationships`：rels 关系文件的生成与管理
- `SharedStrings`：共享字符串 XML 序列化
- `StyleSerializer`：样式与编号格式的 XML 序列化
- `WorksheetXMLGenerator`：单个工作表 XML 生成（友元访问 Worksheet 内部）
- `DocPropsXMLGenerator`：文档属性 XML 生成
- `UnifiedXMLGenerator` 最终负责上述各生成器的注册与编排

