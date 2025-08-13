# 核心层（core）API 与类关系

本节覆盖核心对象：`Workbook`、`Worksheet`、`Cell`、样式系统（`FormatDescriptor`、`FormatRepository`、`StyleBuilder`）、`RangeFormatter`、共享字符串、地址解析等。

## 关系图（简述）
- `Workbook` 聚合：`Worksheet`(N)、`FormatRepository`(1)、`SharedStringTable`(1，可选)、`theme::Theme`(可选)
- `Worksheet` 依赖：父 `Workbook`（弱引用）、`FormatRepository`、`SharedStringTable`、`SharedFormulaManager`、`CellRangeManager`
- `Cell` 由 `Worksheet` 管理；其格式指向不可变 `FormatDescriptor`（可被仓储去重并共享）
- `StyleBuilder` 负责以流式 API 构造 `FormatDescriptor`；`FormatRepository` 负责去重与 ID 分配
- `RangeFormatter` 批量作用于 `Worksheet` 的单元格范围，内部复用 `FormatDescriptor/StyleBuilder`

---

## class fastexcel::core::Workbook
- 职责：表示一个 Excel 工作簿，管理工作表、样式仓储、共享字符串、主题与保存/打开流程；支持只读/编辑两种打开模式与保真写回策略。
- 关键关系：
  - 持有 `std::vector<std::shared_ptr<Worksheet>>` 工作表集合
  - 持有 `std::unique_ptr<FormatRepository>` 样式仓储（ID 0 为默认样式）
  - 可持有 `std::unique_ptr<SharedStringTable>`、`std::unique_ptr<theme::Theme>`
  - 友元：`opc::PackageEditor`、`reader::XLSXReader`、`xml::DocPropsXMLGenerator`
- 主要 API：
  - 生命周期/IO：`create(path)`、`openForReading(path)`、`openForEditing(path)`、`save()`、`saveAs(filename)`、`close()`、`isOpen()`
  - 模式与保真：`setPreserveUnknownParts(bool)`、`getPreserveUnknownParts()`
  - 工作表管理：`addSheet(name)`、`insertSheet(index,name)`、`removeSheet(name/index)`、`getSheet(name/index)`（含 const 版本）、`operator[]`（索引/名称）、`getSheetCount()`、`getSheetNames()`、`hasSheet(name)`、`findSheet(name)`、`getAllSheets()`、`clearAllSheets()`、`renameSheet(old,new)`
  - 单元格便捷：`getValue<T>("Sheet!A1")`、`setValue<T>("Sheet!A1", value)`、`tryGetValue<T>(sheet,row,col)`、`trySetValue<T>(sheet,row,col,value)`
  - 样式：`addStyle(const FormatDescriptor&)`、`addStyle(const StyleBuilder&)`、`addNamedStyle(...)`、`createStyleBuilder()`、`getStyle(id)`、`getDefaultStyleId()`、`isValidStyleId(id)`、`getStyleCount()`、`getStyles()`、`copyStylesFrom(other)`、`getStyleStats()`
  - 主题：`setThemeXML(xml)`、`setOriginalThemeXML(xml)`、`setTheme(theme)`、`getTheme()`、`setThemeName(name)`、`setThemeColor(type,color)`、`setThemeColorByName(name,color)`、`setThemeMajorFont*`/`setThemeMinorFont*`、`getThemeXML()`
  - 批处理/工具：`mergeWorkbook(other, options)`、`exportWorksheets(names, out)`、`batchRenameWorksheets(map)`、`batchRemoveWorksheets(names)`、`reorderWorksheets(order)`
  - 查找替换：`findAndReplaceAll(find, replace, options)`、`findAll(text, options)`
  - 统计与优化：`getStatistics()`、`isModified()`、`getTotalMemoryUsage()`、`optimize()`

使用建议：只读处理用 `openForReading`；编辑保存用 `openForEditing` 或 `create`。

---

## class fastexcel::core::Worksheet
- 职责：表示 Excel 工作表，管理单元格数据、行列/合并/筛选/打印/视图等设置；提供模板化数据访问与范围格式化。
- 关键关系：
  - 拥有 `std::map<(row,col), Cell>` 数据；弱引用父 `Workbook`
  - 使用 `SharedStringTable`、`FormatRepository` 进行字符串/样式复用
  - 协作 `SharedFormulaManager`、`CellRangeManager` 进行公式与使用范围管理
  - 与 `xml::WorksheetXMLGenerator` 为友元（序列化）
- 主要 API（选摘常用）：
  - 值读写（模板/地址）：`getValue<T>(row,col)`、`setValue<T>(row,col,val)`、`getValue<T>("A1")`、`setValue<T>("A1",val)`、`tryGetValue<T>(row,col)`、`getValueOr<T>(row,col,def)`
  - 日期/链接：`writeDateTime(row,col,tm)`、`writeUrl(row,col,url, text)`
  - 单元格格式：`setCellFormat(row,col,const FormatDescriptor&)`、`setCellFormat(row,col,std::shared_ptr<const FormatDescriptor>)`、`setCellFormat(row,col,const StyleBuilder&)`、`tryGetCellFormat(row,col)`
  - 范围格式化：`rangeFormatter("A1:C10")`、`rangeFormatter(sr,sc,er,ec)` 返回 `RangeFormatter`
  - 行列/外观：列宽/行高/隐藏/大纲层级、网格线与标题显示、右到左、选中范围/活动单元格、默认行高列宽等（见源码注释：`showGridlines`、`showRowColHeaders`、`setRightToLeft`、`setActiveCell`、`setSelection` 等）
  - 合并/筛选/冻结/打印：`mergeRange`（见实现）、`setAutoFilter`、`freezePanes`、打印区域/重复行列/页边距/缩放等访问器
  - 查询统计：`getUsedRange()`、`isEmpty()`、`hasData()`、`getRowCount()`、`getColumnCount()`、`getCellCountInRow/Column()`、`hasCellAt(row,col)`、`getCellCount()`
  - 基本信息：`getName()/setName()`、`getSheetId()`、`getParentWorkbook()`

---

## class fastexcel::core::Cell
- 职责：表示单元格及其值/类型/公式/格式/超链接/批注；提供模板化的强类型访问与安全访问封装。
- 关键关系：
  - 由 `Worksheet` 管理；格式指向不可变 `FormatDescriptor`
  - 与 `xml::WorksheetXMLGenerator` 为友元（序列化）
- 主要 API：
  - 类型与值：`CellType getType()`、`template<typename T> T getValue()`、`setValue(const T&)`、`tryGetValue<T>()`、`getValueOr<T>(def)`、`asString/asNumber/asBool/asInt`
  - 公式：`setFormula(str, result)`、共享公式 `setSharedFormula(idx, result)`、`setSharedFormulaReference(idx)`、`getFormula()`、`getFormulaResult()`、`isSharedFormula()`
  - 格式：`setFormat(std::shared_ptr<const FormatDescriptor>)`、`getFormatDescriptor()`、`hasFormat()`
  - 超链接/批注：`setHyperlink(url)`、`getHyperlink()`、`hasHyperlink()`、`setComment(str)`、`getComment()`、`hasComment()`

---

## class fastexcel::core::FormatDescriptor
- 职责：不可变格式值对象，承载字体/对齐/边框/填充/数字格式/保护等属性；可哈希、可比较，便于仓储去重与跨线程共享。
- 关键关系：
  - 由 `StyleBuilder` 构造；`FormatRepository` 负责去重存储并分配 ID
  - 被 `Cell/Worksheet/RangeFormatter` 引用共享
- 主要 API：
  - 获取器：字体、对齐、边框、填充、数字格式、保护等 `get*()` 全量只读访问（见头文件）
  - 能力检查：`hasFont()`、`hasFill()`、`hasBorder()`、`hasAlignment()`、`hasProtection()`、`hasAnyFormatting()`
  - 值对象：`operator==`、`hash()`、`getDefault()`、`modify()`（产生修改版 builder）

---

## class fastexcel::core::FormatRepository
- 职责：线程安全的格式仓储，负责 `FormatDescriptor` 去重、ID 分配与查询；提供统计与内存估算。
- 关键关系：
  - `Workbook` 持有一个仓储实例；`Worksheet/RangeFormatter` 通过它共享格式
- 主要 API：
  - 存取：`addFormat(const FormatDescriptor&) -> id`、`getFormat(id)`、`getDefaultFormat()`、`getDefaultFormatId()`、`isValidFormatId(id)`、`getFormatCount()`、`clear()`
  - 统计：`getCacheHitRate()`、`getDeduplicationStats()`、`getMemoryUsage()`
  - 跨工作簿：`importFormats(source_repo, id_mapping)`

---

## class fastexcel::core::StyleBuilder
- 职责：以链式/流式 API 构建 `FormatDescriptor`；支持从已有 `FormatDescriptor` 派生修改。
- 关键关系：
  - 输出 `FormatDescriptor`，供 `FormatRepository/Worksheet/RangeFormatter` 使用
- 主要 API（链式）：
  - 字体：`fontName/Size/Color/bold/italic/underline/strikeout/superscript/subscript`
  - 对齐：`horizontalAlign/verticalAlign/leftAlign/centerAlign/rightAlign/vcenterAlign/textWrap/rotation`
  - 边框：各边与颜色、对角线类型（见头文件）
  - 填充：`backgroundColor/fg_color/pattern`
  - 数字格式：`num_format/num_format_index`
  - 保护：`locked/hidden`

---

## class fastexcel::core::RangeFormatter
- 职责：对工作表的单元格范围批量应用样式，支持表格样式、边框快捷方法，延迟到 `apply()` 统一执行。
- 关键关系：
  - 绑定 `Worksheet`；内部复用/生成 `FormatDescriptor` 并通过 `FormatRepository` 共享
- 主要 API：
  - 范围：`setRange(sr,sc,er,ec)`、`setRange("A1:C10")`、`setRow(row, startCol, endCol)`、`setColumn(col, startRow, endRow)`
  - 格式：`applyFormat(desc)`、`applyStyle(builder)`、`applySharedFormat(ptr)`
  - 表格样式：`asTable(style)`、`withHeaders(bool)`、`withBanding(rowBand,colBand)`
  - 边框：`allBorders(style,color)`、`outsideBorders(style,color)`、`insideBorders(style,color)`、`noBorders()`
  - 快捷：`backgroundColor(color)`、`fontColor(color)`、`bold()`、`align(hAlign, vAlign)`、`centerAlign()`、`rightAlign()`
  - 执行：`apply() -> int`（返回受影响单元格数）

---

## 其他核心类型（概览）
- `core::SharedStringTable`：共享字符串表，供 `Worksheet/Cell` 写入字符串时复用索引，减少重复存储。
- `core::CellRangeManager`：跟踪工作表使用范围与批量操作边界。
- `core::SharedFormula/SharedFormulaManager`：共享公式簇与引用管理。
- `core::WorkbookModeSelector`：根据规模/内存阈值自动选择流式/批量模式。
- `core::IFileWriter/StreamingFileWriter/BatchFileWriter`：统一的文件写入抽象与两种实现（流式/批量），供 XML 生成层输出。
- `core::ExcelStructureGenerator`：根据 Workbook 构建标准 XLSX 目录与部件结构的生成器。
- `core::DirtyManager`：统一的“脏状态”管理，供 XML 生成选择性输出。
- `core::QuickFormat`：常用格式快捷封装（如居中/加粗/日期等）。
- `core::StyleTransferContext`：跨工作簿样式复制的 ID 映射上下文。
- `core::Path`：跨平台路径封装；被 `archive/opc/reader` 等广泛使用。
