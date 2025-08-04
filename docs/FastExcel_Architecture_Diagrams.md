# FastExcel 架构流程图

## 1. 整体架构图

```mermaid
graph TB
    subgraph "应用层"
        A1[Basic Usage Example]
        A2[Formatting Example]
        A3[Reader Example]
    end
    
    subgraph "API层"
        B1[FastExcel.hpp - 统一入口]
    end
    
    subgraph "核心层"
        C1[Workbook - 工作簿管理]
        C2[Worksheet - 工作表管理]
        C3[Cell - 单元格数据]
        C4[Format - 格式管理]
        C5[Color - 颜色管理]
    end
    
    subgraph "读取层"
        D1[XLSXReader - 主读取器]
        D2[SharedStringsParser - 共享字符串解析]
        D3[WorksheetParser - 工作表解析]
        D4[StylesParser - 样式解析]
    end
    
    subgraph "XML层"
        E1[XMLStreamWriter - XML写入器]
        E2[ContentTypes - 内容类型]
        E3[Relationships - 关系管理]
        E4[SharedStrings - 共享字符串]
    end
    
    subgraph "归档层"
        F1[FileManager - 文件管理]
        F2[ZipArchive - ZIP处理]
    end
    
    subgraph "工具层"
        G1[Logger - 日志系统]
    end
    
    A1 --> B1
    A2 --> B1
    A3 --> B1
    B1 --> C1
    C1 --> C2
    C2 --> C3
    C2 --> C4
    C4 --> C5
    
    D1 --> D2
    D1 --> D3
    D1 --> D4
    D1 --> F2
    
    C1 --> E1
    C1 --> E2
    C1 --> E3
    C1 --> E4
    
    E1 --> F1
    F1 --> F2
    
    C1 --> G1
    D1 --> G1
```

## 2. 写入数据流程图

```mermaid
sequenceDiagram
    participant User as 用户代码
    participant WB as Workbook
    participant WS as Worksheet
    participant Cell as Cell
    participant Format as Format
    participant XML as XMLStreamWriter
    participant ZIP as ZipArchive
    participant File as 文件系统
    
    User->>WB: 创建工作簿
    User->>WS: 添加工作表
    User->>WS: 写入数据
    WS->>Cell: 创建单元格
    User->>Format: 设置格式
    Cell->>Format: 应用格式
    
    User->>WB: 保存文件
    WB->>XML: 生成XML内容
    XML->>ZIP: 写入ZIP文件
    ZIP->>File: 保存到磁盘
    
    Note over WB,ZIP: 流式处理，内存友好
```

## 3. 读取数据流程图

```mermaid
sequenceDiagram
    participant User as 用户代码
    participant Reader as XLSXReader
    participant ZIP as ZipArchive
    participant SSP as SharedStringsParser
    participant WSP as WorksheetParser
    participant WB as Workbook
    participant WS as Worksheet
    
    User->>Reader: 打开XLSX文件
    Reader->>ZIP: 解压ZIP文件
    ZIP->>Reader: 返回XML文件列表
    
    Reader->>SSP: 解析共享字符串
    SSP->>Reader: 返回字符串表
    
    Reader->>WSP: 解析工作表
    WSP->>Reader: 返回工作表数据
    
    User->>Reader: 加载工作簿
    Reader->>WB: 重建Workbook对象
    WB->>WS: 创建Worksheet对象
    Reader->>User: 返回工作簿
```

## 4. 内存优化架构图

```mermaid
graph LR
    subgraph "Cell内存优化"
        A1[位域标志 - 8bit]
        A2[Union数据 - 16byte]
        A3[扩展数据指针 - 8byte]
        A4[延迟分配扩展数据]
    end
    
    subgraph "共享资源"
        B1[共享字符串表]
        B2[格式池]
        B3[颜色池]
    end
    
    subgraph "缓存机制"
        C1[行缓存]
        C2[XML缓冲区]
        C3[ZIP缓冲区]
    end
    
    A1 --> A2
    A2 --> A3
    A3 --> A4
    
    B1 --> C1
    B2 --> C1
    C1 --> C2
    C2 --> C3
```

## 5. 性能优化策略图

```mermaid
mindmap
  root((性能优化))
    内存优化
      Cell优化
        位域压缩
        Union设计
        延迟分配
      共享资源
        字符串表
        格式池
      缓存机制
        行缓存
        XML缓冲
    I/O优化
      流式写入
        固定缓冲区
        直接文件写入
      批量操作
        批量ZIP写入
        批量格式设置
    并发优化
      线程安全
        互斥锁
        原子操作
      锁优化
        细粒度锁
        读写锁
```

## 6. 模块依赖关系图

```mermaid
graph TD
    subgraph "第三方依赖"
        T1[minizip-ng]
        T2[zlib-ng]
        T3[libexpat]
        T4[fmt]
        T5[googletest]
    end
    
    subgraph "FastExcel模块"
        M1[Archive Layer]
        M2[XML Layer]
        M3[Core Layer]
        M4[Reader Layer]
        M5[Utils Layer]
    end
    
    T1 --> M1
    T2 --> M1
    T3 --> M2
    T4 --> M5
    T5 --> M5
    
    M1 --> M3
    M2 --> M3
    M5 --> M3
    M1 --> M4
    M2 --> M4
    M3 --> M4
```

## 7. 类关系图

```mermaid
classDiagram
    class Workbook {
        -worksheets_: vector~Worksheet~
        -formats_: map~int, Format~
        -file_manager_: FileManager
        +addWorksheet(): Worksheet
        +createFormat(): Format
        +save(): bool
    }
    
    class Worksheet {
        -cells_: map~pair~int,int~, Cell~
        -parent_workbook_: Workbook
        +writeString(row, col, value)
        +writeNumber(row, col, value)
        +getCell(row, col): Cell
    }
    
    class Cell {
        -flags_: bitfield
        -value_: union
        -extended_: ExtendedData*
        +setValue(value)
        +setFormat(format)
        +getType(): CellType
    }
    
    class Format {
        -font_properties_
        -alignment_properties_
        -border_properties_
        +setBold(bool)
        +setFontColor(Color)
        +setAlignment(align)
    }
    
    class XLSXReader {
        -zip_archive_: ZipArchive
        -shared_strings_: map
        +open(): bool
        +loadWorksheet(name): Worksheet
        +getMetadata(): WorkbookMetadata
    }
    
    Workbook ||--o{ Worksheet : contains
    Worksheet ||--o{ Cell : contains
    Cell }o--|| Format : uses
    XLSXReader ..> Workbook : creates
```

## 8. 数据处理流水线

```mermaid
flowchart LR
    subgraph "输入阶段"
        I1[用户数据]
        I2[格式设置]
    end
    
    subgraph "处理阶段"
        P1[数据验证]
        P2[格式应用]
        P3[共享字符串处理]
        P4[内存优化]
    end
    
    subgraph "输出阶段"
        O1[XML生成]
        O2[ZIP压缩]
        O3[文件写入]
    end
    
    I1 --> P1
    I2 --> P2
    P1 --> P3
    P2 --> P3
    P3 --> P4
    P4 --> O1
    O1 --> O2
    O2 --> O3
```

## 9. 错误处理流程图

```mermaid
flowchart TD
    A[操作开始] --> B{输入验证}
    B -->|失败| C[记录错误日志]
    B -->|成功| D[执行操作]
    D --> E{操作结果}
    E -->|成功| F[返回结果]
    E -->|失败| G[异常处理]
    G --> H[资源清理]
    H --> I[错误报告]
    C --> I
    I --> J[操作结束]
    F --> J
    
    style C fill:#ffcccc
    style G fill:#ffcccc
    style I fill:#ffcccc
```

## 10. 配置系统架构

```mermaid
graph TB
    subgraph "配置层次"
        C1[全局配置]
        C2[工作簿配置]
        C3[工作表配置]
        C4[单元格配置]
    end
    
    subgraph "配置类型"
        T1[性能配置]
        T2[内存配置]
        T3[功能配置]
        T4[兼容性配置]
    end
    
    C1 --> C2
    C2 --> C3
    C3 --> C4
    
    T1 --> C1
    T2 --> C1
    T3 --> C2
    T4 --> C2
```

这些流程图从不同角度展示了FastExcel的架构设计：

1. **整体架构图**: 展示各层之间的关系
2. **数据流程图**: 展示写入和读取的完整流程
3. **内存优化图**: 展示内存优化策略
4. **性能优化图**: 展示各种性能优化手段
5. **依赖关系图**: 展示模块间的依赖关系
6. **类关系图**: 展示核心类的关系
7. **处理流水线**: 展示数据处理的各个阶段
8. **错误处理图**: 展示错误处理机制
9. **配置系统图**: 展示配置的层次结构

这些图表有助于理解FastExcel的整体设计思路和实现细节。