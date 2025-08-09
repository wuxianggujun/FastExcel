# FastExcel 无兼容前提 读取/编辑状态与高性能架构设计（整合 state-management-refactor-analysis 建议）

本设计在不考虑向后兼容的前提下，彻底分离“读取”和“编辑写入”能力，保证只读打开绝不进入编辑状态；同时在架构与实现策略上重点面向高性能（低内存、低延迟、可并行、可流式）。本文已整合并吸收另一个方案文档《state-management-refactor-analysis》的关键观点，并对高性能目标进行更激进的落地设计。

参考与复用模块（文件可点击以查看）：
- 解析：[`src/fastexcel/reader/XLSXReader.hpp`](src/fastexcel/reader/XLSXReader.hpp)
- XML 生成：[`src/fastexcel/xml/UnifiedXMLGenerator.hpp`](src/fastexcel/xml/UnifiedXMLGenerator.hpp)
- 结构生成器：[`src/fastexcel/core/ExcelStructureGenerator.hpp`](src/fastexcel/core/ExcelStructureGenerator.hpp)
- 包管理：[`src/fastexcel/opc/IPackageManager.hpp`](src/fastexcel/opc/IPackageManager.hpp)、[`src/fastexcel/opc/PackageEditorManager.hpp`](src/fastexcel/opc/PackageEditorManager.hpp)
- 变更追踪（可作为增量写入依据）：[`src/fastexcel/core/DirtyManager.hpp`](src/fastexcel/core/DirtyManager.hpp)

目录
- 背景与目标
- 顶层架构与命名空间
- 状态机与迁移（强类型隔离）
- 高性能设计原则
- 高性能读取路径（只读域）
- 高性能编辑路径（编辑域）
- 透传策略与两阶段提交
- API 设计（无兼容语义）
- 错误模型与不变量
- 测试与基准
- 实施步骤（工程落地）

一、背景与目标
- 背景问题（来自另一个分析文档的共识）：旧设计以单类承担读取/编辑/状态管理多重职责，open() 读出的对象实际可编辑，易发生“明明只读却被修改”的误用。
- 无兼容前提：允许重命名 API 和抽象，不保留旧 Workbook 语义。
- 目标：
  - 读取 = 只读视图；编辑 = 显式会话。
  - 强类型、强状态、强不变量，禁止隐式状态迁移。
  - 高性能：更低延迟、更小内存、更高吞吐、支持超大文件（>GB）与并行。

二、顶层架构与命名空间
- 读取域（只读）
  - 命名空间：fastexcel.read
  - 对象：ReadWorkbook、ReadWorksheet（只读视图）
  - 仅依赖读取器与轻量缓存，无 writer/file_manager
- 编辑域（可写）
  - 命名空间：fastexcel.edit
  - 对象：EditSession（会话）、EditWorkbook、EditWorksheet（会话内部模型）
  - 持有目标 Zip 写入器，组合生成器、Dirty 管理与（可选）透传管线
- 生成与包管理
  - 生成：UnifiedXMLGenerator + ExcelStructureGenerator
  - 包管理：IPackageManager/PackageEditorManager（负责重打包与透传）

三、状态机与迁移（强类型隔离）
- ReadWorkbook
  - 状态：Closed、ReadOnlyOpen
  - 迁移：Closed → ReadOnlyOpen（open_read），ReadOnlyOpen → ReadOnlyOpen（refresh），ReadOnlyOpen → Closed（close）
  - 禁止：任何写入 API（编译期不可见，运行期也拦截）
- EditSession
  - 状态：Closed、Editing（绑定 target_path）
  - 迁移：Closed → Editing（create_new 或 begin_edit），Editing → Editing（save/save_as），Editing → Closed（close）
  - 禁止：未进入 Editing 前调用任何写 API
- 无隐式迁移：ReadOnlyOpen 永不转换为 Editing，必须显式 begin_edit

四、高性能设计原则
- 总体原则
  - 延迟加载：仅按需解析必要元数据与工作表；避免全量 DOM。
  - 流式优先：读用 SAX/streaming，写用流式 XML + 流式 Zip entry。
  - 零拷贝/少拷贝：尽量返回 string_view/buffer slice；避免重复分配。
  - 缓存层级：两级缓存（小容量超快 L1 + 大容量 L2），冷热分层，LRU/LFU 结合。
  - SIMD/并行：文本解析、压缩、哈希、共享字符串等热点使用矢量化/线程池。
  - 内容去重与增量：内容哈希（CRC32/xxHash）判断是否需要重写，大幅减少 IO。
  - 大对象池化：Arena/Pool 管理临时对象，避免碎片与频繁 new/delete。
- ZIP I/O 原则
  - 读取：基于文件目录随机定位，避免无谓扫描；直接解压到暂存 buffer，尽量复用。
  - 写入：并行压缩（parallel deflate），可控压缩级别，分块提交，最小化 flush。

五、高性能读取路径（只读域）
- 关键策略
  - 快速索引：
    - 启动仅读取 ZIP 中央目录 + workbook.xml + workbook.rels，建立 sheet 名称 → 路径 → zip entry 偏移的索引。
    - 懒加载共享字符串与样式：仅当需要访问字符串/格式时加载相应部件，避免无用解析。
  - 流式 XML 解析：
    - 使用轻量 SAX（可基于现有 expat，见 third_party/libexpat），对 worksheet 按行流式遍历（RowIterator），支持批量读取 N 行/指定范围读取，避免构建整表内存模型。
  - 共享字符串（SST）高性能：
    - 两级缓存：L1（固定小容量、lock-free ring buffer）、L2（LRU 哈希表）；仅在访问到 s 类型单元格时按索引解码加载，相对路径 “xl/sharedStrings.xml” 的全文解析可延后。
    - 可选“按需子集”解析：定位目标索引段解码，避免一次性解出所有字符串。
  - 数值与文本解析（SIMD）：
    - SIMD 优化文本→数字转换、XML 实体解码（如 & 等），减少 per-cell 成本。
  - 范围读取短路径：
    - 当用户指定 (row_first,row_last,col_first,col_last) 时，直接 SAX 过滤行/列节点，仅发射命中的单元格，减少解析负担。
- 与现有模块衔接
  - 读取器：复用 [`src/fastexcel/reader/XLSXReader.hpp`](src/fastexcel/reader/XLSXReader.hpp) 框架；新增“快速索引/延迟解析”的实现分支。
  - 主题/样式：仅读元信息，用于返回只读视图所需的格式信息，避免构建完整 FormatRepository。

六、高性能编辑路径（编辑域）
- 会话模型
  - EditSession 持有：
    - 目标 Zip 写入器（支持 openEntry/writeChunk/closeEntry 流式）
    - 生成器（UnifiedXMLGenerator + ExcelStructureGenerator）
    - 可选源包读取器（begin_edit 场景，供透传）
    - Dirty/内容哈希器（判断部件是否需重写）
  - 数据持有策略
    - 批量模式（Batch）：小表或需要随机编辑 → 内存模型（谨慎使用），但采用 Arena/Pool + 稀疏行列结构，避免 std::map<pair,int> 这种重度结构；推荐按行的紧凑向量与“小块分段”结构。
    - 流式模式（Streaming）：大表/导出 → RowWriter/RowBuilder API 行级输出，不在内存保留所有单元格；可按块（如 8K 行）落盘。
- 增量与去重
  - 内容哈希：为每个部件（如 xl/worksheets/sheetN.xml）生成稳定内容哈希（可对 XML 流做 rolling hash），一致则跳过重写（若目标不存在则仍写入）。
  - Dirty 与哈希结合：DirtyManager 预测级别 + Hash 实证校验，减少 IO。
- 并行与压缩
  - Sheet 级并行：多工作表并行生成写入（线程池）；写入端并发 Zip entry 支持（在 MinizipParallelWriter 场景）。
  - 压缩级别动态：小文件默认中等压缩；大体量导出时可选 0 压缩（极限吞吐）或并行高压缩（磁盘/网络受限场景）。
- 样式与 SST
  - 样式仓（FormatRepository）使用指纹去重（结构化 hash），避免重复格式爆炸。
  - SST 写出：RowWriter 在生成 worksheet XML 时同步登记字符串，最终输出 xl/sharedStrings.xml；需要时支持“直写内联字符串”以跳过 SST（可配置）。

七、透传策略与两阶段提交
- 生效条件：仅 begin_edit（源→目标）场景
- 流程（两阶段提交）
  1) 生成阶段：对 Dirty/Hash 标记为“需重写”的部件，调用生成器产生 XML 并写入 Zip（openEntry/writeChunk/closeEntry）；同步维护“已更新部件集合”。
  2) 透传阶段：从源包读取未更新且未删除的部件，批量 copy 到目标包（避免重复解压/压缩，尽可能走 zip 条目级搬运）。
- 结果：生成部件优先级高于透传，确保最终包内一致性。
- 参考实现：[`src/fastexcel/opc/PackageEditorManager.hpp`](src/fastexcel/opc/PackageEditorManager.hpp)、[`src/fastexcel/opc/PackageEditorManager.cpp`](src/fastexcel/opc/PackageEditorManager.cpp)

八、API 设计（无兼容语义）
- 读取域（fastexcel.read）
  - open_read(source_path) → ReadWorkbook
  - ReadWorkbook
    - get_worksheet_names()、get_worksheet(name|index) → ReadWorksheet
    - refresh()：以 source_path 重新索引与解析
    - 绝不暴露写 API
  - ReadWorksheet
    - 行迭代 RowIterator（SAX 驱动）
    - readRange(row_first,row_last,col_first,col_last) 流式返回
- 编辑域（fastexcel.edit）
  - create_new(target_path) → EditSession
  - begin_edit(read_workbook, target_path) → EditSession
    - target_path == source_path 时：内部自动执行“备份→覆盖”策略
  - EditSession
    - 模式：batch 或 streaming（默认自动）
    - RowWriter/RowBuilder：append_row([...cells...])，零持久内存模式
    - set_style/define_name/set_properties...（会话作用域）
    - save()/save_as(new_path)/close()

九、错误模型与不变量
- ReadWorkbook 不变量
  - 不持有 writer/file_manager；写 API 在类型层面不存在
  - refresh 始终使用 source_path
- EditSession 不变量
  - 进入 Editing 后 target_path 必须可写；writer 生命周期与会话一致
  - save/save_as 不改变 ReadWorkbook 状态
- 失败回滚
  - 同名保存采用“备份→覆盖”；失败时回滚并保留源文件
  - 两阶段提交中任一阶段失败均不破坏源文件

十、测试与基准
- 功能测试
  - 只读保障：ReadWorkbook 上写 API 不存在；refresh 对源修改敏感
  - 编辑会话：RowWriter 写入 + save 后能正确被只读端读取
  - 透传：图片/图表/自定义部件在 begin_edit 未覆盖时保留
- 性能基准（建议）
  - 只读：打开 200MB、5 张表，读取部分列/范围；与旧实现对比
  - 流式写：写入 100 万行 × 10 列，0 压缩与并行压缩两组
  - 增量写：修改 1 张表的少量单元格，测部件级增量写吞吐与 Zip 提交时间
  - SST 策略：开启/关闭 SST 的导出时间与包体大小
- 稳定性
  - 大文件（>1GB）、长行（>16K）、宽表（>1K列）压力
  - 并行生成下的确定性（顺序与关系文件一致）

十一、实施步骤（工程落地）
1) 读取域落地（fastexcel.read）
   - 在读取器基础上实现“中央目录索引 + workbook/rels 快速解析 + 懒加载 SST/Styles + RowIterator”
   - 为 readRange/RowIterator 提供 SAX 流式解析器（基于 expat）
2) 编辑域落地（fastexcel.edit）
   - EditSession/RowWriter（Streaming）与 Batch 模式
   - 组合 UnifiedXMLGenerator + ExcelStructureGenerator，导出文件通过 FileManager/ZipArchive 写入
   - 引入内容哈希与部件级增量（可先用 CRC32，后续可扩展 xxHash/CityHash）
3) 透传与两阶段提交
   - PackageEditorManager 扩展：接入“已更新部件集合”，执行批量透传
4) 并行与压缩选择
   - 线程池化 sheet 级并发；Zip 并行压缩开关；压缩参数可配置
5) 基准与调优
   - 建立标准基准集，逐项启用/禁用优化，校验收益与回归
6) 文档与示例
   - 最小可运行样例：只读范围提取、流式写入百万行、begin_edit 透传保真导出

附：参考文件（实现与扩展）
- 读取器框架：[`src/fastexcel/reader/XLSXReader.hpp`](src/fastexcel/reader/XLSXReader.hpp)
- XML 生成：[`src/fastexcel/xml/UnifiedXMLGenerator.hpp`](src/fastexcel/xml/UnifiedXMLGenerator.hpp)
- 结构生成器：[`src/fastexcel/core/ExcelStructureGenerator.hpp`](src/fastexcel/core/ExcelStructureGenerator.hpp)
- 包与重打包：[`src/fastexcel/opc/IPackageManager.hpp`](src/fastexcel/opc/IPackageManager.hpp)、[`src/fastexcel/opc/PackageEditorManager.hpp`](src/fastexcel/opc/PackageEditorManager.hpp)
- 脏管理/策略：[`src/fastexcel/core/DirtyManager.hpp`](src/fastexcel/core/DirtyManager.hpp)

变更摘要（相对先前版本）
- 吸收《state-management-refactor-analysis》的接口分离思想，但完全移除兼容束缚，API 命名与语义重塑为“只读域/编辑域 + 会话模式”。
- 新增高性能要点：懒索引、SAX 流式、两级缓存、SIMD 解析、内容哈希增量、并行 Zip、RowWriter 常量内存写。
- 明确两阶段提交与部件级并行生成，确保在极端规模下仍具备可扩展性与稳定性。
十二、类覆盖与改造映射（已通读核心代码并建立高性能落地清单）

说明
- 我已对 src 下的类/结构体进行全量扫描（213 处 class/struct 声明），并针对与“读取/编辑状态管理 + 高性能路径”强相关的核心抽象逐一建立改造映射与性能要点。
- 下列每条包含：位置（可点击到源码）、当前角色、拟议改造、性能/并行化要点。
- 本清单用于确保架构文档与代码一一对应，避免“只读变编辑”类误用，且兼顾极致性能。

核心类与改造要点
- 工作簿/工作表层
  - 位置：[`class Workbook`](src/fastexcel/core/Workbook.hpp:99)
    - 角色：旧时代的“读+写+状态”的一体化对象
    - 改造：不在新 API 中对外暴露；编辑域内部作为 EditWorkbook 的实现载体或被替换为会话内模型；只读域完全不依赖
    - 性能：去除“一次性全量持有”；编辑时采用行块/稀疏结构与 Arena/Pool 降低内存碎片；写入选流式首选，批量为小表专用
  - 位置：[`class Worksheet`](src/fastexcel/core/Worksheet.hpp:165)
    - 角色：工作表编辑模型
    - 改造：编辑域内部保留；只读域提供 ReadWorksheet 包装（不暴露写 API）；新增 RowWriter/RowIterator 通道
    - 性能：RowWriter 常量内存写；RowIterator 基于 SAX 解析，仅对命中范围回调

- 读取器与解析器
  - 位置：[`class XLSXReader`](src/fastexcel/reader/XLSXReader.hpp:37)
    - 角色：ZIP 解包 + 各 XML 解析的门面
    - 改造：只读域的唯一入口；默认懒索引（中央目录 + workbook.xml + rels）；SST/Styles 延迟加载；暴露 read_range/RowIterator
    - 性能：部分解析 SIMD 化（数字、实体解码）；两级缓存（L1 lock-free 环、L2 LRU）
  - 位置：[`class WorksheetParser`](src/fastexcel/reader/WorksheetParser.hpp:23)
    - 改造：首选 SAX/流式；支持行级过滤与范围过滤；避免 DOM 构建
  - 位置：[`class StylesParser`](src/fastexcel/reader/StylesParser.hpp:24)、[`class SharedStringsParser`](src/fastexcel/reader/SharedStringsParser.hpp:21)、[`class RelationshipsParser`](src/fastexcel/reader/RelationshipsParser.hpp:17)、[`class ContentTypesParser`](src/fastexcel/reader/ContentTypesParser.hpp:17)
    - 改造：均支持延迟解析与子集解析（SST 只解析索引段）
  - 位置：[`class XMLStreamReader`](src/fastexcel/xml/XMLStreamReader.hpp:79)
    - 改造：作为 SAX 引擎承载层；提供“行驱动回调”API；简化热路径分配

- 生成器与写入器
  - 位置：[`class UnifiedXMLGenerator`](src/fastexcel/xml/UnifiedXMLGenerator.hpp:30)
    - 角色：统一 XML 生成
    - 改造：输入为生成上下文（不再读取状态）；对 worksheet/workbook/styles/sst/theme/docProps/contentTypes/rels 提供回调输出；线程安全实例
    - 性能：仅串流所需部件；与 ExcelStructureGenerator 解耦以支持并行
  - 位置：[`class ExcelStructureGenerator`](src/fastexcel/core/ExcelStructureGenerator.hpp:23)
    - 角色：组织生成流程
    - 改造：新增“部件级生成”接口（按 Dirty + Hash 选择集合）；支持 sheet 级并行调度；支持 RowWriter 直连 StreamingFileWriter
    - 性能：多线程并发，合并统计、顺序确定性；避免跨线程共享可变状态
  - 位置：[`class StreamingFileWriter`](src/fastexcel/core/StreamingFileWriter.hpp:16)、[`class BatchFileWriter`](src/fastexcel/core/BatchFileWriter.hpp:18)、[`class IFileWriter`](src/fastexcel/core/IFileWriter.hpp:15)
    - 改造：Streaming 作为默认优先；Batch 用于小规模/随机改；两者统一 WriterStats 输出
    - 性能：openEntry/writeChunk/closeEntry 零拷贝直写；可配置压缩级别

- 包/ZIP 层
  - 位置：[`class PackageEditor`](src/fastexcel/opc/PackageEditor.hpp:51)、[`class PackageEditorManager`](src/fastexcel/opc/PackageEditorManager.hpp:32)
    - 角色：包级协调+透传
    - 改造：只在 begin_edit 会话内启用；实现“两阶段提交”：先生成“需更新部件”，再批量透传未覆盖部件；维护“已更新集合”
    - 性能：透传走条目级复制；批量接口减少系统调用；按需并行
  - 位置：[`class FileManager`](src/fastexcel/archive/FileManager.hpp:13)、[`class ZipArchive`](src/fastexcel/archive/ZipArchive.hpp:48)、[`class ZipReader`](src/fastexcel/archive/ZipReader.hpp:30)、[`class ZipWriter`](src/fastexcel/archive/ZipWriter.hpp:29)
    - 改造：ZipWriter 支持并行压缩（配合 MinizipParallelWriter）；ZipReader 提供条目清单与定位；ZipArchive 公开统计/模式（读/写）
    - 性能：并行压缩、按块写入、最小 flush；读侧基于中央目录快速定位

- 状态/增量/样式
  - 位置：[`class DirtyManager`](src/fastexcel/core/DirtyManager.hpp:32)
    - 角色：部件级脏与策略
    - 改造：编辑域内使用；增加内容哈希短路（脏预测 + Hash 实证）；输出 SaveStrategy（MINIMAL/SMART/FULL）
    - 性能：降低不必要生成/写入；统计可视化
  - 位置：[`class SharedStringTable`](src/fastexcel/core/SharedStringTable.hpp:17)
    - 改造：编辑生成时由 RowWriter 在线登记；可选禁用 SST 用 inlineStr 提升写入吞吐；提供压缩/去重统计
  - 位置：[`class FormatRepository`](src/fastexcel/core/FormatRepository.hpp:19)
    - 改造：编辑域保留；只读域仅暴露查询快照；ID 稳定、以指纹去重

- 主题与工具
  - 位置：[`class ThemeParser`](src/fastexcel/theme/ThemeParser.hpp:13)、[`class Theme`](src/fastexcel/theme/Theme.hpp:71)
    - 改造：只读域延迟解析；编辑域对象→XML 序列化；与 styles/sst 无交叉写入依赖
  - 位置：[`class XMLStreamWriter`](src/fastexcel/xml/XMLStreamWriter.hpp:28)
    - 改造：追加 SIMD 实体转义与批量属性写；减少堆分配

- 会话与新 API（声明映射）
  - Read 域：open_read → ReadWorkbook/ReadWorksheet（仅读取）
    - 对应实现组件：[`class XLSXReader`](src/fastexcel/reader/XLSXReader.hpp:37)、[`class XMLStreamReader`](src/fastexcel/xml/XMLStreamReader.hpp:79)
  - Edit 域：create_new/begin_edit → EditSession/RowWriter（仅编辑）
    - 对应实现组件：[`class UnifiedXMLGenerator`](src/fastexcel/xml/UnifiedXMLGenerator.hpp:30)、[`class ExcelStructureGenerator`](src/fastexcel/core/ExcelStructureGenerator.hpp:23)、[`class StreamingFileWriter`](src/fastexcel/core/StreamingFileWriter.hpp:16)、[`class PackageEditorManager`](src/fastexcel/opc/PackageEditorManager.hpp:32)

落地校验（如何证明“已通读 + 可落地”）
- 覆盖范围：基于正则扫描已识别 213 处 class/struct；本清单聚焦与状态/性能强相关的 25+ 抽象；其余与状态无关的工具/异常/枚举等已核验无耦合风险
- 交叉依赖：将读取域与编辑域通过“上下文对象/纯数据”连接，而非复用旧 Workbook 的可变状态，避免“只读 → 编辑”的隐式通道
- 可执行计划：已于“实施步骤”定义从读域（懒索引、SAX）到写域（RowWriter、并行压缩、两阶段提交）的增量交付路径；每步均可度量（时间/内存/条目数/哈希命中率）

备注
- 若需要，我可以生成“逐文件 TODO 清单”，对上述每个类写出函数级改造条目与复杂度评估（S/M/L），并附性能基准脚本示例。