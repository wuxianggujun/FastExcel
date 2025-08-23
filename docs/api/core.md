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
  - **生命周期/IO**：
    - `create(filepath)` - 创建新Excel文件
    - `openReadOnly(filepath)` - 只读方式打开Excel文件
    - `openEditable(filepath)` - 编辑方式打开Excel文件  
    - `save()` - 保存工作簿
    - `saveAs(filename)` - 另存为
  - **CSV集成**：
    - `loadCSV(filepath, sheet_name, options)` - 从CSV文件创建工作表
    - `loadCSVString(csv_content, sheet_name, options)` - 从CSV字符串创建工作表
  - **工作表管理**：
    - `addSheet(name)` - 添加新工作表
    - `insertSheet(index, name)` - 在指定位置插入工作表
    - `getSheet(name/index)` - 获取工作表（含const版本）
    - `operator[](index/name)` - 工作表访问操作符
    - `getFirstSheet()` / `getLastSheet()` - 获取首个/最后工作表
    - `findSheet(name)` - 查找工作表
    - `getAllSheets()` - 获取所有工作表
    - `copyWorksheet(source_name, new_name)` - 复制工作表
    - `getActiveWorksheet()` - 获取活动工作表
    - `tryGetSheet(name/index)` - 安全获取工作表（返回optional）
    - `removeSheet(name/index)` - 删除工作表
    - `getSheetCount()` - 获取工作表数量
    - `hasSheet(name)` - 检查工作表是否存在
  - **样式系统**：
    - `createStyleBuilder()` - 创建样式构建器
    - `getFormatRepository()` - 获取格式仓库
  - **工作簿模式**：
    - `setWorkbookMode(mode)` - 设置工作簿模式（AUTO/BATCH/STREAMING）

使用建议：只读处理用 `openForReading`；编辑保存用 `openForEditing` 或 `create`。

---

## class fastexcel::core::Worksheet
- 职责：表示 Excel 工作表，管理单元格数据、样式、图片、CSV集成等功能；提供现代化的数据访问接口。
- 关键关系：
  - 拥有单元格数据管理；弱引用父 `Workbook`
  - 使用 `SharedStringTable`、`FormatRepository` 进行字符串/样式复用
  - 集成 `CellDataProcessor`、`WorksheetCSVHandler` 等管理器
  - 与 `xml::WorksheetXMLGenerator` 协作进行XML序列化
- 主要 API：
  - **单元格访问**：
    - `getCell(row, col)` / `getCell("A1")` / `getCell(Address)` - 获取单元格引用
    - `getValue<T>(row, col)` / `getValue<T>("A1")` - 泛型值获取
    - `setValue<T>(row, col, value)` / `setValue<T>("A1", value)` - 泛型值设置
    - `tryGetValue<T>()` / `getValueOr<T>()` - 安全访问方法
  - **公式支持**：
    - `setFormula(row, col, formula, result)` - 设置单元格公式
    - `setFormula(Address, formula, result)` - 使用地址设置公式
  - **批量操作**：
    - `writeRange(start_addr, data)` - 批量写入数据范围
    - `writeRow(row, start_col, data)` - 写入行数据
    - `writeColumn(col, start_row, data)` - 写入列数据
    - `setRangeFormat(range, format)` - 批量设置范围格式
  - **图片操作**：
    - `insertImage(row, col, image_path)` - 插入图片
    - `insertImage(Address, image)` - 使用地址插入图片
    - `insertImage(range, image_path)` - 在范围内插入图片
  - **CSV集成**：
    - `importCSV(file_path, options)` - 导入CSV数据
    - `exportCSV(file_path, options)` - 导出为CSV文件
  - **工作表信息**：
    - `getName()` / `setName(name)` - 工作表名称管理
    - `getUsedRange()` - 获取已使用范围
    - `getCellCount()` - 获取单元格数量
    - `isEmpty()` / `hasData()` - 数据状态查询
  - **链式操作**（通过WorksheetChain）：
    - 支持流畅的链式API调用模式

---

## class fastexcel::core::Cell
- 职责：表示单元格及其值/类型/公式/格式；提供类型安全的数据访问和格式设置。
- 关键关系：
  - 由 `Worksheet` 管理；格式指向不可变 `FormatDescriptor`
  - 与 `xml::WorksheetXMLGenerator` 协作进行XML序列化
  - 支持7种Excel数据类型：数字、字符串、布尔、公式、日期时间、错误值、空值
- 主要 API：
  - **值操作**：
    - `setValue<T>(value)` - 泛型值设置，支持自动类型推导
    - `getValue<T>()` - 泛型值获取，支持类型转换
    - `tryGetValue<T>()` - 安全获取，返回optional
    - `getValueOr<T>(default_value)` - 获取值或返回默认值
  - **类型便捷访问**：
    - `asString()` - 作为字符串获取
    - `asNumber()` - 作为数字获取  
    - `asBool()` - 作为布尔值获取
    - `asInt()` - 作为整数获取
  - **公式支持**：
    - `setFormula(formula, result)` - 设置公式和缓存结果
    - `getFormula()` - 获取公式字符串
    - `hasFormula()` - 检查是否包含公式
  - **格式设置**：
    - `setFormat(FormatDescriptor)` - 设置单元格格式
    - `getFormatDescriptor()` - 获取格式描述符
  - **状态查询**：
    - `isEmpty()` - 检查是否为空
    - `getType()` - 获取数据类型
    - `isModified()` - 检查是否被修改

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
- 职责：通过流畅的链式API构建不可变的`FormatDescriptor`；支持从现有格式创建修改版本。
- 关键关系：
  - 输出 `FormatDescriptor`，供 `FormatRepository/Cell/Worksheet` 使用
  - 隐藏底层格式系统复杂性，提供用户友好的接口
- 主要 API（所有方法都支持链式调用）：
  - **字体设置**：
    - `fontName(name)` - 设置字体名称
    - `fontSize(size)` - 设置字体大小（1.0-409.0）
    - `font(name, size)` - 同时设置名称和大小
    - `font(name, size, bold)` - 设置字体和粗体
    - `bold(is_bold)` - 设置粗体
    - `italic(is_italic)` - 设置斜体
    - `underline(type)` - 设置下划线类型
    - `strikeout(is_strike)` - 设置删除线
    - `fontColor(color)` - 设置字体颜色
    - `script(type)` - 设置上标/下标
  - **对齐方式**：
    - `horizontalAlign(align)` - 水平对齐
    - `verticalAlign(align)` - 垂直对齐
    - `textWrap(wrap)` - 文本换行
    - `rotation(degrees)` - 旋转角度
    - `indent(level)` - 缩进级别
    - `shrinkToFit(shrink)` - 收缩适应
  - **边框设置**：
    - `leftBorder(style, color)` - 左边框
    - `rightBorder(style, color)` - 右边框
    - `topBorder(style, color)` - 上边框
    - `bottomBorder(style, color)` - 下边框
    - `allBorders(style, color)` - 所有边框
    - `diagonalBorder(style, color, type)` - 对角线边框
  - **填充样式**：
    - `backgroundColor(color)` - 背景色
    - `foregroundColor(color)` - 前景色
    - `pattern(type)` - 图案类型
    - `solidFill(color)` - 纯色填充
  - **数字格式**：
    - `numberFormat(format)` - 自定义数字格式
    - `numberFormat(type)` - 内置数字格式类型
  - **保护设置**：
    - `locked(is_locked)` - 单元格锁定
    - `hidden(is_hidden)` - 公式隐藏
  - **构建方法**：
    - `build()` - 生成最终的FormatDescriptor对象

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
