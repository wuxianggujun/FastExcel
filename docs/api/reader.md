# 读取器（reader）API 与类关系

面向 XLSX 文件解析：通过 `archive::ZipArchive` 读取包中文件，解析工作簿/工作表/样式/共享字符串/主题等，填充核心模型。

## 关系图（简述）
- `XLSXReader`：高层读取器
  - 组合：`archive::ZipArchive`
  - 输出：构造并填充 `core::Workbook`、`core::Worksheet`
  - 解析：调用内部解析函数与各子解析器（ContentTypes/Relationships/SharedStrings/Styles/Worksheet 等）
  - 主题：可返回原始主题 XML 与解析后的 `theme::Theme`

---

## class fastexcel::reader::XLSXReader
- 职责：以系统层 API（返回 `core::ErrorCode`）解析 XLSX；支持只读加载、按需加载工作表、获取元数据/定义名/主题等。
- 主要 API：
  - 构造/状态：`XLSXReader(filename|Path)`、`~XLSXReader()`、`open()`、`close()`、`isOpen()`、`getFilename()`
  - 加载：`loadWorkbook(unique_ptr<core::Workbook>&)`、`loadWorksheet(name, shared_ptr<core::Worksheet>&)`、`getSheetNames(vector<string>&)`
  - 元数据：`getMetadata(WorkbookMetadata&)`
  - 定义名称：`getDefinedNames(vector<string>&)`
  - 样式与映射：`getStyles()`（index->FormatDescriptor）、`getStyleIdMapping()`（源样式 ID -> FormatRepository ID）
  - 主题：`getThemeXML()`、`getParsedTheme()`
- 内部解析（概览）：`parseWorkbookXML()`、`parseWorksheetXML(path, ws)`、`parseStylesXML()`、`parseSharedStringsXML()`、`parseContentTypesXML()`、`parseRelationshipsXML()`、`parseDocPropsXML()`、`parseThemeXML()`；若干辅助如 `extractXMLFromZip(path)`、`extractAttribute` 等。

---

## 子解析器（概览）
- `ContentTypesParser`、`RelationshipsParser`、`SharedStringsParser`、`StylesParser`、`WorksheetParser`：面向具体部件的解析工具类（见对应头文件）。

