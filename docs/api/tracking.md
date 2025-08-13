# 变更跟踪（tracking）API 与类关系

用于跟踪 XLSX 包部件的脏标记与变更统计，服务于 `opc::PackageEditor` 的选择性重建与提交。

## 关系图（简述）
- `IChangeTracker`：抽象接口
- `StandardChangeTracker`：默认实现，使用哈希集合管理脏部件；可智能传播关联部件的脏标记
- `ChangeTrackerService`：工厂/门面（如有），供 `opc::PackageEditor` 注入

---

## interface fastexcel::tracking::IChangeTracker
- 职责：最小职责集合的变更跟踪接口。
- 主要 API：`markPartDirty(part)`、`markPartClean(part)`、`isPartDirty(part)`、`getDirtyParts()`、`clearAll()`、`hasChanges()`

## class fastexcel::tracking::StandardChangeTracker
- 职责：`IChangeTracker` 的标准实现，专注部件级脏标记与简单统计，O(1) 查询。
- 主要 API：实现接口的全部方法；内部 `markRelatedDirty(part)` 进行关联传播。

