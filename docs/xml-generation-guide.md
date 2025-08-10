# FastExcel XML 生成统一规范

目的：统一 Excel XML 的生成职责边界、顺序与写入策略，消除重复逻辑，确保批量与流式两种模式下的行为一致且确定。

结论摘要：
- 采用“统一调度 + 分散生成”的架构。
- 由统一调度器 ExcelStructureGenerator 负责顺序与写入；各领域对象仅负责“生成自身 XML 片段”的回调。
- 通过 IFileWriter 策略接口屏蔽批量/流式细节，保证一套流程两种实现。

一、核心参与者与职责
- 统一调度器：ExcelStructureGenerator，负责完整生成流程与顺序控制、调用各片段生成回调、以及写入。
- 写入策略：IFileWriter 抽象，具体实现为 BatchFileWriter 与 StreamingFileWriter。
- 领域对象：Workbook、Worksheet、SharedStringTable 等，仅提供 generate...XML(callback) 的片段输出。

参考代码入口：
- 统一调度 generate 实现：[ExcelStructureGenerator::generate()](src/fastexcel/core/ExcelStructureGenerator.cpp:23)
- 构造与注入策略：[ExcelStructureGenerator::ExcelStructureGenerator()](src/fastexcel/core/ExcelStructureGenerator.cpp:11)
- 获取类型名（便于日志/统计）：[ExcelStructureGenerator::getGeneratorType()](src/fastexcel/core/ExcelStructureGenerator.cpp:77)
- 写入策略接口文件：[src/fastexcel/core/IFileWriter.hpp](src/fastexcel/core/IFileWriter.hpp)
- 批量实现文件：[src/fastexcel/core/BatchFileWriter.hpp](src/fastexcel/core/BatchFileWriter.hpp)
- 流式实现文件：[src/fastexcel/core/StreamingFileWriter.hpp](src/fastexcel/core/StreamingFileWriter.hpp)

二、统一的 XML 生成顺序（严格、确定、可复现）
统一调度器按如下顺序生成所有文件，任何模式下顺序不变：
1) [Content_Types].xml → [Workbook::generateContentTypesXML()](src/fastexcel/core/Workbook.cpp:1239)
2) _rels/.rels → [Workbook::generateRelsXML()](src/fastexcel/core/Workbook.cpp:1311)
3) docProps/app.xml → [Workbook::generateDocPropsAppXML()](src/fastexcel/core/Workbook.cpp:1066)
4) docProps/core.xml → [Workbook::generateDocPropsCoreXML()](src/fastexcel/core/Workbook.cpp:1140)
5) docProps/custom.xml（若存在） → [Workbook::generateDocPropsCustomXML()](src/fastexcel/core/Workbook.cpp:1210)
6) xl/_rels/workbook.xml.rels → [Workbook::generateWorkbookRelsXML()](src/fastexcel/core/Workbook.cpp:1348)
7) xl/workbook.xml → [Workbook::generateWorkbookXML()](src/fastexcel/core/Workbook.cpp:973)
8) xl/styles.xml → [Workbook::generateStylesXML()](src/fastexcel/core/Workbook.cpp:1029)
9) xl/theme/theme1.xml → [Workbook::generateThemeXML()](src/fastexcel/core/Workbook.cpp:1393)
10) xl/worksheets/sheetN.xml（按 N 从 1 递增）→ [Workbook::generateWorksheetXML()](src/fastexcel/core/Workbook.cpp:1062)
11) xl/worksheets/_rels/sheetN.xml.rels（可选，若该表有链接等）
12) xl/sharedStrings.xml（当启用共享字符串时生成，允许为空）→ [Workbook::generateSharedStringsXML()](src/fastexcel/core/Workbook.cpp:1034)

说明：
- 顺序与 libxlsxwriter 对齐，确保 Excel 一致性。
- 即使 sharedStrings 为空，但启用时也要生成空文件，且在 Content_Types 与 workbook.xml.rels 中必须包含条目/关系。

三、关系文件与 rId 分配规范
- _rels/.rels：固定为 rId1→xl/workbook.xml，rId2→docProps/core.xml，rId3→docProps/app.xml，rId4→docProps/custom.xml（若有）。
- xl/_rels/workbook.xml.rels：
  - 先为每个工作表按顺序分配 rId，从 rId1 开始依次指向 worksheets/sheetN.xml。
  - 之后依次添加 theme/theme1.xml、styles.xml。
  - 若启用共享字符串，追加 sharedStrings.xml，rId 继续顺延，保证稳定、确定。

四、共享字符串（SST）统一策略
- 数据来源：Worksheet 在生成单元格时将字符串通过 Workbook::addSharedString 累积到 SST；SST 仅存储唯一值并返回 id。
- 生成时机：调度器在所有工作表生成完成后，最后生成 xl/sharedStrings.xml。
- 空表处理：启用共享字符串但无内容时，生成合法的空 sst 文档（count=0, uniqueCount=0）。
- 内存统计：SST 内存由 SharedStringTable 负责统计与清空，避免 Workbook 内部再维护冗余的 shared_strings_* 容器。

五、批量 vs 流式 的统一
- 通过 IFileWriter 屏蔽差异：
  - 批量：一次性收集字符串并写入（BatchFileWriter）。
  - 流式：openStreamingFile → 持续 writeStreamingChunk → closeStreamingFile（StreamingFileWriter）。
- 选择策略：ExcelStructureGenerator 采用“智能阈值 + 显式模式”：
  - 显式模式：WorkbookOptions.mode 为 BATCH/STREAMING 时强制。
  - AUTO 模式：依据数据量与内存评估在 Worksheet 级别做决策（小表批量，大表流式）。

六、统一的 API 合约（片段生成回调）
- 领域对象仅提供"生成自身 XML 内容"的回调接口，写入由调度器负责：
  - Workbook：generateContentTypesXML、generateRelsXML、generateDocPropsAppXML、generateDocPropsCoreXML、generateDocPropsCustomXML、generateWorkbookRelsXML、generateWorkbookXML、generateStylesXML、generateThemeXML、generateSharedStringsXML（均为 callback 方式，见上方链接）。
  - Worksheet：generateXML、generateRelsXML（回调写入，接口见 [src/fastexcel/core/Worksheet.hpp](src/fastexcel/core/Worksheet.hpp)）。
  - SharedStringTable：generateXML（回调写入，见 [src/fastexcel/core/SharedStringTable.hpp](src/fastexcel/core/SharedStringTable.hpp)）。

**重要说明**：
- 所有领域对象的 XML 生成方法都应该使用回调函数 `std::function<void(const char*, size_t)>` 作为参数
- 不应该有直接写文件的方法（如 `generateXMLToFile`），这些是旧架构的遗留，应标记为废弃
- 写入策略（批量/流式）完全由 IFileWriter 实现决定，领域对象不需要关心

七、调用流程（简化）
1) 构造 IFileWriter（按模式选择批量或流式）。
2) 构造 ExcelStructureGenerator 并注入 writer 与 workbook。
3) 调用 [ExcelStructureGenerator::generate()](src/fastexcel/core/ExcelStructureGenerator.cpp:23)：
   - generateBasicFiles：基础部件（含 sharedStrings 条目但内容最后生成）。
   - generateWorksheets：按智能阈值为每个表选择 batch/streaming，统一通过 Worksheet::generateXML 回调输出。
   - finalize：批量模式 flush。
   - 写入统计记录。

八、顺序与一致性“硬性约束”（可用于断言/单测）
- A1. 所有 sheetN.xml 必须在 sharedStrings.xml 之前写入。
- A2. Content_Types.xml 必须包含 sharedStrings 覆盖项（当启用共享字符串时），且 workbook.xml.rels 必须包含 sharedStrings 关系。
- A3. workbook.xml 中的 sheets 顺序需与 sheetN.xml 的 N 一致，r:id 与 workbook.xml.rels 的 rId 序一致。
- A4. 任何模式下输出字节流的“文件顺序”和“关系顺序”保持一致。
- A5. 主题与样式文件在工作簿之后但在工作表之前生成。

九、迁移与落地建议
- 短期：
  - 保留 Workbook 内已有 generateExcelStructure… 系列方法，对外设置为“兼容层”，其内部改为构造 ExcelStructureGenerator 并调用 generate，避免重复维护。
- 中期：
  - 将写入与顺序控制彻底移出 Workbook，仅保留各 generate...XML 片段回调；新增/修改仅在 ExcelStructureGenerator 与 IFileWriter 上演进。
- 长期：
  - 提供 io::IWorkbookWriter 抽象，形成文件格式插件化（xlsx、ods 等），统一复用 ExcelStructureGenerator 的策略与顺序规则。

十、如何回答“是否要把所有生成方法都放到一个类？”
- 不建议把所有 XML 细节都集中到一个类。原因：
  - 领域知识属于领域对象（如样式、主题、工作簿元数据、工作表网格），放回各自类中更清晰。
  - 统一调度器只关心顺序与写入，避免成为“上帝类”，同时策略切换更简单。
- 因此推荐的结构是：调度器统一顺序 + 对象内聚片段生成 + 策略化写入。

十一、相关实现位置快速索引
- 调度器与策略：
  - [src/fastexcel/core/ExcelStructureGenerator.hpp](src/fastexcel/core/ExcelStructureGenerator.hpp)
  - [src/fastexcel/core/ExcelStructureGenerator.cpp](src/fastexcel/core/ExcelStructureGenerator.cpp)
  - [src/fastexcel/core/IFileWriter.hpp](src/fastexcel/core/IFileWriter.hpp)
  - [src/fastexcel/core/BatchFileWriter.hpp](src/fastexcel/core/BatchFileWriter.hpp)
  - [src/fastexcel/core/StreamingFileWriter.hpp](src/fastexcel/core/StreamingFileWriter.hpp)
- 工作簿片段回调（函数链接见第“统一顺序”章）：
  - [src/fastexcel/core/Workbook.cpp](src/fastexcel/core/Workbook.cpp)
- 工作表回调：
  - [src/fastexcel/core/Worksheet.hpp](src/fastexcel/core/Worksheet.hpp)
- 共享字符串：
  - [src/fastexcel/core/SharedStringTable.hpp](src/fastexcel/core/SharedStringTable.hpp)

十二、后续工作清单（建议）
- 将 Workbook::generateExcelStructureBatch/Streaming 标记为 @deprecated，并在实现内部委托给 ExcelStructureGenerator。
- 将所有领域对象中的 `generateXMLToFile` 等直接写文件方法标记为 @deprecated，统一使用回调方式
- 在单测中增加对“文件输出顺序”“rId 顺序”“sharedStrings 为空/非空”三类情况的断言。
- 在性能文档中统一说明批量/流式的阈值与可调参数，并暴露到 WorkbookOptions。

---

附录：新架构（Orchestrator + Part Generators）补充

- 内容协调层：`xml::UnifiedXMLGenerator`
  - 注册 `IXMLPartGenerator` 实现，提供 `generateParts(writer, parts)` 与 `generateAll(writer)`。
  - 支持与 DirtyManager 协作的过滤式生成：只写需要更新的部件（`generateParts(writer, parts, dirtyManager)`）。
- 部件生成层：`IXMLPartGenerator` 实现类（位于 `src/fastexcel/xml/`）
  - 代表类：ContentTypes/RootRels/WorkbookParts/Styles/SharedStrings/Theme/DocProps/Worksheets/WorksheetRels。
  - 全部采用回调 + IFileWriter，真流式输出，避免大字符串中转。
- 流程/策略层：`core::ExcelStructureGenerator`
  - 通过 `generateParts` 精确调度：基础部件 → 工作表/表关系 → 共享字符串（最后）。
  - 仍由 `Workbook::shouldGenerate*`/DirtyManager 负责增量判定。
