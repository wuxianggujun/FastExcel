# OPC 包编辑（opc）API 与类关系

负责面向 XLSX 的 Open Packaging Conventions（OPC）层面的包编辑、部件依赖管理与变更编排。

## 关系图（简述）
- `PackageEditor` 作为编排者：
  - 依赖 `IPackageManager` 处理包读写与部件管理（通常封装 `archive::ZipReader/Writer`）
  - 依赖 `xml::UnifiedXMLGenerator` 生成 XML 部件
  - 依赖 `tracking::IChangeTracker` 跟踪部件脏标记
  - 与 `core::Workbook` 双向交互（读取状态、提交更改）
- 相关辅助：`PackageEditorManager`、`PartGraph`（部件图）、`ZipRepackWriter`（重打包）、`PackageManagerService`（工厂/服务门面）

---

## class fastexcel::opc::PackageEditor
- 职责：统一封装“打开/从 Workbook 构建/创建新包、检测变更、选择性重建、提交保存”的流程。
- 主要 API：
  - 工厂：`open(Path)`、`fromWorkbook(core::Workbook*)`、`create()`
  - 保存：`save()`（回原路径）、`commit(dst)`（另存）
  - 变更：`detectChanges()`、`markPartDirty(part)`、`getChangeStats()`、`getDirtyParts()`、`isDirty()`
  - 访问：`getWorkbook()`、`getSheetNames()`、`getAllParts()`
  - 配置：`setPreserveUnknownParts(bool)`、`setRemoveCalcChain(bool)`、`setAutoDetectChanges(bool)`、`getOptions()`
  - 验证/调试：`isValidSheetName(name)`、`isValidCellRef(r,c)`、`generatePart(part)`、`validateXML(xml)`
- 关系：内部持有 `unique_ptr<IPackageManager>`、`unique_ptr<xml::UnifiedXMLGenerator>`、`unique_ptr<tracking::IChangeTracker>`；弱持有 `core::Workbook*`。

---

## 接口与实现（概览）
- `IPackageManager`：抽象包管理接口；典型实现组合 `archive::ZipReader/ZipWriter`；被 `PackageManagerService` 创建或注入。
- `ZipRepackWriter`：在保留部分原始压缩流的情况下进行重打包，提高性能并保持保真。
- `PartGraph`：部件依赖图与关系解析，用于变更传播与选择性重建。
- `PackageEditorManager`：编辑器管理/缓存（如果存在）。
- `PackageManagerService`：创建 `IPackageManager` 或提供服务门面。

