# FastExcel 编辑保存架构（脏部件 + 透传）设计与使用说明

本文档说明新版保存架构的核心思想、实现细节、脏标记策略、按需生成规则、验证方法与后续扩展建议。该架构遵循“只改我改过的，其它一律不动”的目标，确保编辑已存在的复杂 Excel 文件时，不破坏图片、批注、绘图、媒体、打印设置、自定义项、VBA 等未修改内容。


## 1. 设计目标

- 编辑模式（基于已有 xlsx 打开）：仅覆盖被实际修改的 OPC 部件，其余全部原样保留（透传）。
- 新建模式（从空白创建）：生成完整必要结构。
- save 与 saveAs 行为一致：两者都先透传原包（若存在），再按脏标记覆盖生成需要更新的部件。
- 不改变工作表数量/顺序/编号时，保持 sheetN.xml 与 sheetN.xml.rels 编号与关系一致，避免破坏图片/批注/超链接等引用。


## 2. 基本流程

保存时整体分两步：

1) 透传阶段
   - 若工作簿来自现有文件且启用“保留未知部件”，则将原包中“全部”条目复制到新包（不跳过任何前缀）。
   - 注意：透传后的条目会在下一步被“按需覆盖”。

2) 覆盖阶段（按脏标记 + 生成规则）
   - 仅对需要更新的部件调用生成器写入并覆盖透传版本；未标记为需要覆盖的部件保持透传不动。

这样实现“编辑了什么就覆盖什么，其它原样透传”。


## 3. 脏标记（DirtyFlags）

在 Workbook 引入 DirtyFlags（部分字段）：
- workbook_core      -> 涉及 xl/workbook.xml 与 xl/_rels/workbook.xml.rels
- root_rels          -> 涉及 _rels/.rels
- content_types      -> 涉及 [Content_Types].xml（仅新增/删除部件时）
- styles             -> 涉及 xl/styles.xml（修改样式时）
- theme              -> 涉及 xl/theme/theme1.xml（修改主题时）
- sst                -> 涉及 xl/sharedStrings.xml（启用共享字符串且内容变化时）
- docprops_core/app/custom -> 文档属性改动时
- sheet_xml[i]       -> 第 i 个工作表的 sheetN.xml 是否需要覆盖
- sheet_rels[i]      -> 第 i 个工作表的 sheetN.xml.rels 是否需要覆盖

在 Worksheet 的编辑 API 内部对脏标记落点：
- 会改变 sheet XML 的操作（写单元格/格式/行列宽高/合并/筛选/冻结/打印设置等） -> 标记 sheet_xml 脏
- 会改变工作表关系的操作（如写超链接） -> 同时标记 sheet_rels 脏

开发者可通过 Workbook 的 markXXXDirty() 接口在必要时显式标脏。


## 4. 生成规则（按需覆盖）

生成器在各阶段调用前，询问 Workbook：
- shouldGenerateContentTypes() -> 决定是否写 [Content_Types].xml
- shouldGenerateRootRels()     -> 是否写 _rels/.rels
- shouldGenerateWorkbookCore() -> 是否写 xl/_rels/workbook.xml.rels 与 xl/workbook.xml
- shouldGenerateStyles()       -> 是否写 xl/styles.xml
- shouldGenerateTheme()        -> 是否写 xl/theme/theme1.xml
- shouldGenerateSharedStrings()-> 是否写 xl/sharedStrings.xml
- shouldGenerateDocPropsCore/App/Custom() -> 是否写 docProps/*
- shouldGenerateSheet(i)       -> 是否写 xl/worksheets/sheet(i+1).xml
- shouldGenerateSheetRels(i)   -> 是否写 xl/worksheets/_rels/sheet(i+1).xml.rels

当处于“透传编辑模式”时：
- 仅当对应 dirty 或确有所需（如主题存在且需覆盖）时才生成；否则跳过，保留透传版本。
- 新建模式/非透传模式下，按旧逻辑完全生成。


## 5. save / saveAs 行为

- save：
  - 若工作簿为“编辑模式 + 启用透传”，则先透传原包全部条目，再按脏标记覆盖。
  - 否则（新建模式），完全生成必要结构。

- saveAs：
  - 打开目标文件（新包），延续原“编辑来源”信息，执行与 save 相同的“透传 + 覆盖”流程。


## 6. 关键点与兼容性

- 不跳过任何透传前缀：透传阶段复制全部条目，避免误删 worksheets/_rels/sheetN.xml.rels 等关系文件。
- 编号对齐：不改变工作表数量/顺序/编号时，保持 sheetN.xml 与 .rels 的映射不变，图片/批注/超链接等原有关系不会受影响。
- 共享字符串：启用时在 finalize 阶段按需生成；若未变化，shouldGenerateSharedStrings() 返回 false 则跳过，透传保留。
- 文档属性/主题/样式：在对应 setter 或修改点标记 dirty，使生成器按需覆盖；未修改即透传。


## 7. 验证建议

建议用以下场景回归：

1) 仅编辑单元格：
   - 修改 sheet1 的部分值保存；检查 xl/worksheets/sheet1.xml 更新，其它如 xl/worksheets/_rels/sheet1.xml.rels、xl/drawings/*、xl/media/*、comments、printerSettings 等全部存在且不变。

2) 调整列宽/行高/合并单元格：
   - 验证 sheetN.xml 中 <cols>、行 ht、自定义合并范围 <mergeCells> 是否正确写出；其它部件不变。

3) 写入超链接：
   - 存在超链接时生成 sheetN.xml.rels；未新增时保留透传。

4) 不改任何东西直接保存：
   - 结果包与输入包（除时间戳、压缩顺序等）应基本一致。

5) 复杂文件（含图片/批注/绘图/自定义部件/VBA）：
   - 仅编辑数据后保存，确保所有非编辑部件依然可见有效。


## 8. 后续扩展（重命名/重排/增删工作表）

若后续需要支持：
- 改名/重排/增删工作表 -> 需要：
  - 更新 xl/workbook.xml 与 xl/_rels/workbook.xml.rels 的 rId 与目标关系
  - 更新 xl/worksheets/_rels/sheetN.xml.rels 的编号或目标
  - 更新 [Content_Types].xml 的工作表 Override 列表
  - 设置 dirty_workbook_core 与 dirty_content_types，必要时对应 sheet_rels 也标脏

可在此架构上再加一层“编号重映射器”，确保所有关系一致性。


## 9. 常见问题（FAQ）

- 问：为什么透传时不跳过任何路径？
  - 答：先透传“全部”，再按脏标记“按需覆盖”。这样可最大程度避免误删重要关系与部件（例如 worksheets/_rels）。

- 问：生成器跳过某些文件会不会导致缺失？
  - 答：不会，因为透传已经把原文件复制过来；跳过仅表示“保持原样，不覆盖”。

- 问：共享字符串是否会被错误覆盖？
  - 答：shouldGenerateSharedStrings() 控制写入，仅当启用且需要更新时才生成；否则保持透传。


## 10. 开发者指南（落点与扩展）

- 在 Worksheet 的编辑 API 内部打点 markSheetXmlDirty/markSheetRelsDirty；新增会影响关系的特性时记得对 sheet_rels 标脏。
- 在修改样式/主题/文档属性时，调用对应 markStylesDirty/markThemeDirty/markDocPropsCore/App/CustomDirty。
- 新增/删除部件时，记得 markContentTypesDirty，并按需设置 workbook_core。
- 透传开关：Workbook::setPreserveUnknownParts(true)（默认已启用）。
- 编辑来源：通过 Workbook::open(path) 读取现有文件将自动标记 opened_from_existing_ 与 original_package_path_。


## 11. 目录与文件影响（摘要）

- 透传：
  - 原包所有条目（包括 [Content_Types].xml、_rels/*、docProps/*、xl/**/*、自定义目录等）
- 按需覆盖（示例）：
  - xl/worksheets/sheetN.xml                  （sheet_xml[i] 为 true）
  - xl/worksheets/_rels/sheetN.xml.rels       （sheet_rels[i] 为 true 且生成 XML 非空）
  - xl/workbook.xml、xl/_rels/workbook.xml.rels（workbook_core 为 true）
  - xl/styles.xml（styles 为 true）
  - xl/theme/theme1.xml（theme 为 true）
  - xl/sharedStrings.xml（sst 为 true）
  - docProps/core.xml、docProps/app.xml、docProps/custom.xml（对应 dirty 为 true）
  - [Content_Types].xml（content_types 为 true）


## 12. 结语

通过“透传 + 按需覆盖”的保存架构，FastExcel 在编辑已有复杂 Excel 文件时可最大限度地保持未修改部件的完整与一致，真正做到“只改我改过的”。如需扩展到工作表的增删改名与重排，只需要在当前框架上增加编号重映射与相应的脏标记设置即可。

