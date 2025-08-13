# FastExcel 新API 完整使用指南

## 概述

本文档提供 FastExcel 新API 的完整使用指南，展示 Workbook 和 Worksheet 如何协同工作，为用户提供简洁、高效、类型安全的 Excel 操作体验。

## 核心设计思想

### 职责分离原则

```
📁 Workbook (工作簿级别)
├── 📄 文件操作: create(), open(), save()
├── 📋 工作表管理: sheet(), addWorksheet(), removeWorksheet()
├── 🎨 全局样式: addStyle(), createStyleBuilder()
├── 🔍 跨表操作: findAll(), replaceAll()
└── ⚙️ 工作簿配置: setDocumentProperties(), protect()

📊 Worksheet (工作表级别)  
├── 🔢 单元格操作: getValue<T>(), setValue<T>()
├── 📐 范围操作: getRange<T>(), setRange<T>()
├── 📏 行列管理: setRowHeight(), setColumnWidth()
├── 🔗 链式调用: set().set().set()
└── 📝 工作表配置: setName(), protect(), freezePanes()
```

### API 层级关系

```cpp
// ✅ 正确的层级关系
auto workbook = Workbook::open("file.xlsx");        // 1. 工作簿级别
auto worksheet = workbook->sheet("数据表");          // 2. 获取工作表
std::string value = worksheet->getValue<std::string>(1, 1); // 3. 工作表操作单元格

// ❌ 错误的设计（原始版本的问题）
// std::string value = workbook->getValue<std::string>(0, 1, 1); // 违反职责分离
```

## 完整 API 参考

### Workbook API 完整列表

#### 1. 工厂方法
```cpp
class Workbook {
public:
    // 创建和打开
    static std::unique_ptr<Workbook> create(const Path& path);
    static std::unique_ptr<Workbook> open(const Path& path);
    static std::unique_ptr<Workbook> openForReading(const Path& path);
    static std::unique_ptr<Workbook> openForEditing(const Path& path);
    static std::unique_ptr<Workbook> createWithTemplate(const Path& path, const Path& template_path);
    
    // 文件操作
    bool save();
    bool saveAs(const std::string& filename);
    bool close();
    bool isOpen() const;
};
```

#### 2. 工作表管理
```cpp
    // 工作表访问（核心API）
    std::shared_ptr<Worksheet> sheet(size_t index);
    std::shared_ptr<Worksheet> sheet(const std::string& name);
    std::shared_ptr<Worksheet> operator[](size_t index);
    std::shared_ptr<Worksheet> operator[](const std::string& name);
    
    // 工作表管理
    std::shared_ptr<Worksheet> addWorksheet(const std::string& name = "");
    std::shared_ptr<Worksheet> insertWorksheet(size_t index, const std::string& name = "");
    bool removeWorksheet(const std::string& name);
    bool removeWorksheet(size_t index);
    
    // 工作表查询
    size_t getWorksheetCount() const;
    std::vector<std::string> getWorksheetNames() const;
    bool hasWorksheet(const std::string& name) const;
    
    // 工作表操作
    bool renameWorksheet(const std::string& old_name, const std::string& new_name);
    bool moveWorksheet(size_t from_index, size_t to_index);
    std::shared_ptr<Worksheet> copyWorksheet(const std::string& source_name, const std::string& new_name);
```

#### 3. 全局操作
```cpp
    // 跨工作表搜索
    struct FindResult {
        std::string worksheet_name;
        int row, col;
        std::string found_text;
        CellType cell_type;
    };
    
    std::vector<FindResult> findAll(const std::string& search_text, const FindOptions& options = {}) const;
    int replaceAll(const std::string& find_text, const std::string& replace_text, const FindOptions& options = {});
    
    // 跨工作表数据获取
    template<typename T>
    std::map<std::string, T> getFromAllSheets(const std::string& address) const;
    
    // 批量工作表操作
    int batchRenameWorksheets(const std::vector<BatchRenameRule>& rules);
    int batchRemoveWorksheets(const std::vector<std::string>& worksheet_names);
    bool reorderWorksheets(const std::vector<std::string>& new_order);
```

#### 4. 样式管理
```cpp
    // 样式仓储
    int addStyle(const FormatDescriptor& style);
    int addStyle(const StyleBuilder& builder);
    std::shared_ptr<const FormatDescriptor> getStyle(int style_id) const;
    
    // 样式构建
    StyleBuilder createStyleBuilder() const;
    
    // 命名样式
    int addNamedStyle(const std::string& name, const FormatDescriptor& style);
    int getNamedStyleId(const std::string& name) const;
    
    // 主题管理
    void setTheme(const Theme& theme);
    const Theme* getTheme() const;
    
    // 样式统计和优化
    StyleStats getStyleStatistics() const;
    size_t optimizeStyles();
```

#### 5. 文档属性
```cpp
    // 核心属性
    void setTitle(const std::string& title);
    void setAuthor(const std::string& author);
    void setSubject(const std::string& subject);
    void setCompany(const std::string& company);
    
    // 批量设置
    void setDocumentProperties(const std::string& title, const std::string& author = "",
                              const std::string& subject = "", const std::string& company = "");
    
    // 自定义属性
    void setCustomProperty(const std::string& name, const std::string& value);
    void setCustomProperty(const std::string& name, double value);
    void setCustomProperty(const std::string& name, bool value);
    std::string getCustomProperty(const std::string& name) const;
    
    // 保护
    void protect(const std::string& password = "", bool lock_structure = true);
    void unprotect(const std::string& password = "");
    bool isProtected() const;
```

#### 6. 性能控制
```cpp
    // 工作模式
    enum class Mode { AUTO, BATCH, STREAMING, INTERACTIVE };
    void setMode(Mode mode);
    Mode getMode() const;
    
    // 内存管理
    struct MemoryStats { size_t total_memory, worksheet_memory, style_memory, string_table_memory, cache_memory; };
    MemoryStats getMemoryStatistics() const;
    void setMemoryLimit(size_t max_bytes);
    void optimizeMemory();
    
    // 缓存控制
    void setCachePolicy(CachePolicy policy);
    void clearCache();
    void warmupCache(const std::vector<std::string>& worksheet_names);
```

### Worksheet API 完整列表

#### 1. 单元格访问
```cpp
class Worksheet {
public:
    // 核心访问API
    template<typename T> T getValue(int row, int col) const;
    template<typename T> void setValue(int row, int col, const T& value);
    
    // Excel地址格式
    template<typename T> T getValue(const std::string& address) const;
    template<typename T> void setValue(const std::string& address, const T& value);
    
    // 安全访问
    template<typename T> std::optional<T> tryGetValue(int row, int col) const noexcept;
    template<typename T> T getValueOr(int row, int col, const T& default_value) const noexcept;
    
    // 底层单元格访问
    Cell& getCell(int row, int col);
    const Cell& getCell(int row, int col) const;
};
```

#### 2. 范围操作
```cpp
    // 一维范围操作
    template<typename T> std::vector<T> getRange(const std::string& range) const;
    template<typename T> void setRange(const std::string& range, const std::vector<T>& values);
    
    // 二维数组支持
    template<typename T> std::vector<std::vector<T>> getRangeAs2D(const std::string& range) const;
    template<typename T> void setRangeFrom2D(const std::string& range, const std::vector<std::vector<T>>& data);
    
    // 流式处理
    template<typename T, typename Processor>
    void processRange(const std::string& range, Processor processor) const;
    
    // 行列批量操作
    template<typename T> std::vector<T> getRow(int row, int start_col = 0, int end_col = -1) const;
    template<typename T> std::vector<T> getColumn(int col, int start_row = 0, int end_row = -1) const;
    template<typename T> void setRow(int row, const std::vector<T>& values, int start_col = 0);
    template<typename T> void setColumn(int col, const std::vector<T>& values, int start_row = 0);
```

#### 3. 链式调用
```cpp
    // 链式设置
    template<typename T> Worksheet& set(int row, int col, const T& value);
    template<typename T> Worksheet& set(const std::string& address, const T& value);
    template<typename T> Worksheet& setRow(int row, const std::vector<T>& values, int start_col = 0);
    template<typename T> Worksheet& setColumn(int col, const std::vector<T>& values, int start_row = 0);
    
    // 链式格式化
    Worksheet& format(int row, int col, int style_id);
    Worksheet& format(const std::string& range, int style_id);
    
    // 链式结构设置
    Worksheet& rowHeight(int row, double height);
    Worksheet& columnWidth(int col, double width);
    Worksheet& merge(const std::string& range);
    Worksheet& freeze(int row, int col);
```

#### 4. 行列管理
```cpp
    // 行操作
    void setRowHeight(int row, double height);
    double getRowHeight(int row) const;
    void hideRow(int row, bool hidden = true);
    bool isRowHidden(int row) const;
    void groupRows(int start_row, int end_row, int outline_level = 1);
    
    // 列操作
    void setColumnWidth(int col, double width);
    double getColumnWidth(int col) const;
    void hideColumn(int col, bool hidden = true);
    bool isColumnHidden(int col) const;
    void groupColumns(int start_col, int end_col, int outline_level = 1);
    
    // 自动调整
    void autoFitRowHeight(int row);
    void autoFitColumnWidth(int col);
    void autoFitAllRows();
    void autoFitAllColumns();
    
    // 插入删除
    void insertRows(int row, int count = 1);
    void insertColumns(int col, int count = 1);
    void deleteRows(int row, int count = 1);
    void deleteColumns(int col, int count = 1);
    
    // 批量行列操作
    void setRowHeights(int start_row, int end_row, double height);
    void setColumnWidths(int start_col, int end_col, double width);
```

#### 5. 工作表属性
```cpp
    // 基本属性
    void setName(const std::string& name);
    std::string getName() const;
    void setVisible(bool visible);
    bool isVisible() const;
    
    // 保护
    struct ProtectionOptions { /* ... */ };
    void protect(const std::string& password = "", const ProtectionOptions& options = {});
    void unprotect(const std::string& password = "");
    bool isProtected() const;
    
    // 视图设置
    struct ViewOptions { bool show_gridlines; bool show_row_column_headers; int zoom_scale; /* ... */ };
    void setViewOptions(const ViewOptions& options);
    ViewOptions getViewOptions() const;
    
    // 冻结窗格
    void freezePanes(int row, int col);
    void unfreezePages();
    bool hasFrozenPanes() const;
    
    // 打印设置
    struct PrintOptions { bool landscape; double scale; /* ... */ };
    void setPrintOptions(const PrintOptions& options);
    void setPrintArea(const std::string& range);
    void setRepeatRows(int first_row, int last_row);
    void setRepeatColumns(int first_col, int last_col);
```

#### 6. 高级功能
```cpp
    // 合并单元格
    void mergeCells(const std::string& range);
    void unmergeCells(const std::string& range);
    bool isMergedCell(int row, int col) const;
    std::vector<std::string> getAllMergedRanges() const;
    
    // 超链接
    void setHyperlink(int row, int col, const std::string& url, const std::string& display_text = "");
    std::string getHyperlink(int row, int col) const;
    void removeHyperlink(int row, int col);
    
    // 批注
    void setComment(int row, int col, const std::string& comment, const std::string& author = "");
    std::string getComment(int row, int col) const;
    void removeComment(int row, int col);
    
    // 数据验证
    struct ValidationRule { ValidationType type; std::string formula1, formula2; /* ... */ };
    void setDataValidation(const std::string& range, const ValidationRule& rule);
    ValidationRule getDataValidation(int row, int col) const;
    
    // 自动筛选
    void setAutoFilter(const std::string& range);
    void removeAutoFilter();
    std::string getAutoFilterRange() const;
    
    // 条件格式
    struct ConditionalFormat { Type type; std::string formula; int style_id; };
    void addConditionalFormat(const std::string& range, const ConditionalFormat& format);
    void removeConditionalFormats(const std::string& range);
    
    // 查询方法
    struct UsedRange { int first_row, last_row, first_col, last_col; bool isEmpty() const; };
    UsedRange getUsedRange() const;
    int getMaxRow() const;
    int getMaxColumn() const;
    size_t getCellCount() const;
```

## 实际使用场景

### 1. 基础数据操作

```cpp
// 🚀 简单的数据读写
void basicDataOperation() {
    auto wb = Workbook::open("data.xlsx");
    auto ws = wb->sheet(0);
    
    // 读取数据 - 类型安全
    std::string product_name = ws->getValue<std::string>("A1");
    double price = ws->getValue<double>("B1");
    int quantity = ws->getValue<int>("C1");
    bool in_stock = ws->getValue<bool>("D1");
    
    // 写入数据 - 自动类型转换
    ws->setValue("A2", "新产品");      // 字符串
    ws->setValue("B2", 299.99);       // 浮点数
    ws->setValue("C2", 50);           // 整数
    ws->setValue("D2", true);         // 布尔值
    
    // 批量操作
    auto all_prices = ws->getRange<double>("B1:B100");
    double total = std::accumulate(all_prices.begin(), all_prices.end(), 0.0);
    
    wb->save();
}
```

### 2. 表格创建和格式化

```cpp
// 🚀 创建格式化的数据表
void createFormattedTable() {
    auto wb = Workbook::create("formatted_table.xlsx");
    auto ws = wb->addWorksheet("产品列表");
    
    // 定义样式
    auto header_style = wb->createStyleBuilder()
        .font().bold(true).size(12).color(Color::WHITE)
        .fill().pattern(FillPattern::SOLID).color(Color::DARK_BLUE)
        .border().all(BorderStyle::THIN)
        .alignment().horizontal(HorizontalAlignment::CENTER)
        .build();
    int header_style_id = wb->addStyle(header_style);
    
    auto data_style = wb->createStyleBuilder()
        .border().all(BorderStyle::THIN)
        .build();
    int data_style_id = wb->addStyle(data_style);
    
    // 创建表头并格式化
    ws->setRow(0, std::vector<std::string>{
        "产品编码", "产品名称", "类别", "单价", "库存", "备注"
    }).format("A1:F1", header_style_id)
      .rowHeight(0, 25.0);
    
    // 设置列宽
    ws->setColumnWidths(0, 5, {12.0, 25.0, 15.0, 10.0, 8.0, 20.0});
    
    // 添加示例数据
    std::vector<ProductInfo> products = getProductData();
    for (size_t i = 0; i < products.size(); ++i) {
        int row = i + 1;
        ws->set(row, 0, products[i].code)
          .set(row, 1, products[i].name)
          .set(row, 2, products[i].category)
          .set(row, 3, products[i].price)
          .set(row, 4, products[i].stock)
          .set(row, 5, products[i].note);
    }
    
    // 格式化数据区域
    std::string data_range = "A2:F" + std::to_string(products.size() + 1);
    ws->format(data_range, data_style_id);
    
    // 添加高级功能
    ws->setAutoFilter("A1:F" + std::to_string(products.size() + 1));
    ws->freezePanes(1, 0);  // 冻结表头行
    
    // 数据验证：库存必须是正整数
    ValidationRule stock_validation;
    stock_validation.type = ValidationType::WHOLE;
    stock_validation.formula1 = "0";
    stock_validation.formula2 = "10000";
    stock_validation.error_message = "库存数量必须是0-10000之间的整数";
    ws->setDataValidation("E2:E1000", stock_validation);
    
    wb->save();
}
```

### 3. 数据分析和统计

```cpp
// 🚀 数据分析示例
void performDataAnalysis() {
    auto wb = Workbook::openForReading("sales_data.xlsx");
    auto ws = wb->sheet("销售数据");
    
    // 获取数据范围
    auto used_range = ws->getUsedRange();
    std::cout << "数据范围: " << used_range.first_row << "," << used_range.first_col 
              << " 到 " << used_range.last_row << "," << used_range.last_col << std::endl;
    
    // 批量读取分析数据
    auto dates = ws->getRange<std::string>("A2:A" + std::to_string(used_range.last_row + 1));
    auto amounts = ws->getRange<double>("C2:C" + std::to_string(used_range.last_row + 1));
    auto regions = ws->getRange<std::string>("D2:D" + std::to_string(used_range.last_row + 1));
    
    // 按地区统计
    std::map<std::string, double> region_totals;
    std::map<std::string, int> region_counts;
    
    for (size_t i = 0; i < amounts.size() && i < regions.size(); ++i) {
        const std::string& region = regions[i];
        double amount = amounts[i];
        
        region_totals[region] += amount;
        region_counts[region]++;
    }
    
    // 创建分析结果工作表
    auto analysis_ws = wb->addWorksheet("分析结果");
    
    // 输出地区统计
    analysis_ws->set("A1", "地区统计分析")
               .set("A3", "地区").set("B3", "销售总额").set("C3", "订单数量").set("D3", "平均金额");
    
    int row = 4;
    for (const auto& [region, total] : region_totals) {
        int count = region_counts[region];
        double average = total / count;
        
        analysis_ws->set(row, 0, region)
                   .set(row, 1, total)
                   .set(row, 2, count)
                   .set(row, 3, average);
        row++;
    }
    
    // 计算总体统计
    double grand_total = std::accumulate(amounts.begin(), amounts.end(), 0.0);
    double average_order = grand_total / amounts.size();
    double max_amount = *std::max_element(amounts.begin(), amounts.end());
    double min_amount = *std::min_element(amounts.begin(), amounts.end());
    
    // 输出总体统计
    analysis_ws->set("F3", "总体统计")
               .set("F4", "总销售额:").set("G4", grand_total)
               .set("F5", "订单总数:").set("G5", static_cast<int>(amounts.size()))
               .set("F6", "平均订单:").set("G6", average_order)
               .set("F7", "最大订单:").set("G7", max_amount)
               .set("F8", "最小订单:").set("G8", min_amount);
    
    wb->save();
}
```

### 4. 大数据批量处理

```cpp
// 🚀 高性能大数据处理
void processBigData() {
    // 设置高性能模式
    auto wb = Workbook::open("big_data.xlsx");
    wb->setMode(Workbook::Mode::BATCH);
    wb->setMemoryLimit(512 * 1024 * 1024);  // 512MB 内存限制
    
    auto ws = wb->sheet(0);
    
    // 流式处理大数据集，避免内存溢出
    const int BATCH_SIZE = 1000;
    std::vector<ProcessedData> batch_results;
    batch_results.reserve(BATCH_SIZE);
    
    int processed_count = 0;
    auto used_range = ws->getUsedRange();
    
    for (int start_row = 1; start_row <= used_range.last_row; start_row += BATCH_SIZE) {
        int end_row = std::min(start_row + BATCH_SIZE - 1, used_range.last_row);
        
        // 批量读取
        auto batch_data = ws->getRangeAs2D<std::string>(
            "A" + std::to_string(start_row) + ":E" + std::to_string(end_row));
        
        // 批量处理
        batch_results.clear();
        for (const auto& row : batch_data) {
            if (row.size() >= 5) {
                ProcessedData result;
                result.id = row[0];
                result.name = row[1];
                result.score = std::stod(row[2]);
                result.category = row[3];
                result.processed_value = calculateProcessedValue(result.score);
                
                batch_results.push_back(result);
            }
        }
        
        // 批量写入结果
        auto result_ws = wb->sheet("处理结果");
        for (size_t i = 0; i < batch_results.size(); ++i) {
            int row = processed_count + i + 1;
            const auto& result = batch_results[i];
            
            result_ws->set(row, 0, result.id)
                     .set(row, 1, result.name)
                     .set(row, 2, result.score)
                     .set(row, 3, result.category)
                     .set(row, 4, result.processed_value);
        }
        
        processed_count += batch_results.size();
        
        // 定期释放内存
        if (processed_count % (BATCH_SIZE * 10) == 0) {
            wb->optimizeMemory();
            std::cout << "已处理 " << processed_count << " 条记录" << std::endl;
        }
    }
    
    std::cout << "总计处理 " << processed_count << " 条记录" << std::endl;
    wb->save();
}
```

### 5. 模板驱动报表生成

```cpp
// 🚀 基于模板的报表生成
class ReportGenerator {
private:
    std::unique_ptr<Workbook> template_wb_;
    std::map<std::string, int> template_styles_;
    
public:
    ReportGenerator(const std::string& template_path) {
        template_wb_ = Workbook::open(template_path);
        
        // 缓存模板样式
        cacheTemplateStyles();
    }
    
    std::unique_ptr<Workbook> generateMonthlyReport(const MonthlyData& data) {
        // 1. 创建新报表
        auto report = Workbook::create("monthly_report_" + data.month + ".xlsx");
        
        // 2. 复制模板样式
        auto style_mapping = report->copyStylesFrom(*template_wb_);
        
        // 3. 设置文档属性
        report->setDocumentProperties(
            data.month + "月销售报表",
            "系统自动生成",
            "月度销售数据分析",
            "ABC公司"
        );
        
        // 4. 创建各个报表页
        generateSummarySheet(report.get(), data, style_mapping);
        generateDetailSheet(report.get(), data, style_mapping);
        generateChartSheet(report.get(), data, style_mapping);
        
        return report;
    }

private:
    void generateSummarySheet(Workbook* wb, const MonthlyData& data, const StyleMapping& styles) {
        auto ws = wb->addWorksheet("汇总");
        
        // 标题
        ws->set("A1", data.month + "月销售汇总报表")
          .format("A1", styles.getMappedId("title_style"))
          .merge("A1:F1")
          .rowHeight(0, 30);
        
        // 关键指标
        ws->set("A3", "关键指标")
          .format("A3:F3", styles.getMappedId("header_style"));
        
        ws->set("A4", "总销售额:").set("B4", data.total_sales)
          .set("A5", "订单数量:").set("B5", data.order_count)
          .set("A6", "平均订单:").set("B6", data.average_order)
          .set("A7", "同比增长:").set("B7", data.growth_rate);
        
        // 地区分解
        ws->set("D4", "地区分解")
          .format("D4:F4", styles.getMappedId("header_style"));
        
        int row = 5;
        for (const auto& region_data : data.regions) {
            ws->set(row, 3, region_data.name)
              .set(row, 4, region_data.sales)
              .set(row, 5, region_data.percentage);
            row++;
        }
        
        // 设置数值格式
        ws->format("B4:B7", styles.getMappedId("currency_style"));
        ws->format("E5:E" + std::to_string(row-1), styles.getMappedId("currency_style"));
        ws->format("F5:F" + std::to_string(row-1), styles.getMappedId("percentage_style"));
    }
    
    void generateDetailSheet(Workbook* wb, const MonthlyData& data, const StyleMapping& styles) {
        auto ws = wb->addWorksheet("明细");
        
        // 表头
        std::vector<std::string> headers = {
            "日期", "订单号", "客户", "产品", "数量", "单价", "金额", "地区", "销售员"
        };
        
        ws->setRow(0, headers)
          .format("A1:I1", styles.getMappedId("header_style"))
          .rowHeight(0, 25);
        
        // 设置列宽
        ws->setColumnWidths(0, 8, {12, 15, 20, 25, 8, 10, 12, 10, 15});
        
        // 数据填充
        for (size_t i = 0; i < data.details.size(); ++i) {
            int row = i + 1;
            const auto& detail = data.details[i];
            
            ws->set(row, 0, detail.date)
              .set(row, 1, detail.order_id)
              .set(row, 2, detail.customer)
              .set(row, 3, detail.product)
              .set(row, 4, detail.quantity)
              .set(row, 5, detail.unit_price)
              .set(row, 6, detail.amount)
              .set(row, 7, detail.region)
              .set(row, 8, detail.salesperson);
        }
        
        // 格式化数据区域
        std::string data_range = "A2:I" + std::to_string(data.details.size() + 1);
        ws->format(data_range, styles.getMappedId("data_style"));
        
        // 设置筛选和冻结
        ws->setAutoFilter("A1:I" + std::to_string(data.details.size() + 1));
        ws->freezePanes(1, 0);
    }
};

// 使用示例
void generateReports() {
    ReportGenerator generator("templates/monthly_report_template.xlsx");
    
    MonthlyData january_data = loadMonthlyData("2025-01");
    auto january_report = generator.generateMonthlyReport(january_data);
    january_report->save();
    
    MonthlyData february_data = loadMonthlyData("2025-02");
    auto february_report = generator.generateMonthlyReport(february_data);
    february_report->save();
}
```

### 6. 动态数据管理

```cpp
// 🚀 动态数据管理系统
class DynamicDataManager {
private:
    std::unique_ptr<Workbook> workbook_;
    std::map<std::string, std::shared_ptr<Worksheet>> worksheets_;
    
public:
    DynamicDataManager(const std::string& filename) {
        workbook_ = Workbook::openForEditing(filename);
        
        // 缓存所有工作表
        for (const auto& name : workbook_->getWorksheetNames()) {
            worksheets_[name] = workbook_->sheet(name);
        }
    }
    
    // 动态添加数据表
    void createDataTable(const std::string& table_name, const std::vector<ColumnDefinition>& columns) {
        auto ws = workbook_->addWorksheet(table_name);
        worksheets_[table_name] = ws;
        
        // 创建表头
        std::vector<std::string> headers;
        for (const auto& col : columns) {
            headers.push_back(col.name);
        }
        
        ws->setRow(0, headers)
          .format("A1:" + char('A' + headers.size() - 1) + "1", getHeaderStyleId())
          .rowHeight(0, 25);
        
        // 设置列宽和验证规则
        for (size_t i = 0; i < columns.size(); ++i) {
            ws->setColumnWidth(i, columns[i].width);
            
            if (columns[i].validation.type != ValidationType::NONE) {
                std::string col_range = char('A' + i) + std::string("2:") + 
                                       char('A' + i) + "1000";
                ws->setDataValidation(col_range, columns[i].validation);
            }
        }
    }
    
    // 动态插入记录
    template<typename RecordType>
    void insertRecord(const std::string& table_name, const RecordType& record) {
        auto ws = getWorksheet(table_name);
        if (!ws) return;
        
        int next_row = ws->getUsedRange().last_row + 1;
        insertRecordAtRow(ws.get(), next_row, record);
    }
    
    // 批量插入记录
    template<typename RecordType>
    void insertRecords(const std::string& table_name, const std::vector<RecordType>& records) {
        auto ws = getWorksheet(table_name);
        if (!ws) return;
        
        int start_row = ws->getUsedRange().last_row + 1;
        
        for (size_t i = 0; i < records.size(); ++i) {
            insertRecordAtRow(ws.get(), start_row + i, records[i]);
        }
    }
    
    // 查询数据
    template<typename T>
    std::vector<T> queryRecords(const std::string& table_name, 
                               const QueryCondition& condition) {
        auto ws = getWorksheet(table_name);
        if (!ws) return {};
        
        std::vector<T> results;
        auto used_range = ws->getUsedRange();
        
        for (int row = 1; row <= used_range.last_row; ++row) {
            if (matchesCondition(ws.get(), row, condition)) {
                T record = extractRecord<T>(ws.get(), row);
                results.push_back(record);
            }
        }
        
        return results;
    }
    
    // 更新记录
    template<typename RecordType>
    int updateRecords(const std::string& table_name, 
                     const QueryCondition& condition,
                     const RecordType& new_values) {
        auto ws = getWorksheet(table_name);
        if (!ws) return 0;
        
        int updated_count = 0;
        auto used_range = ws->getUsedRange();
        
        for (int row = 1; row <= used_range.last_row; ++row) {
            if (matchesCondition(ws.get(), row, condition)) {
                updateRecordAtRow(ws.get(), row, new_values);
                updated_count++;
            }
        }
        
        return updated_count;
    }
    
    // 删除记录
    int deleteRecords(const std::string& table_name, const QueryCondition& condition) {
        auto ws = getWorksheet(table_name);
        if (!ws) return 0;
        
        std::vector<int> rows_to_delete;
        auto used_range = ws->getUsedRange();
        
        // 收集要删除的行（从后往前，避免索引变化）
        for (int row = used_range.last_row; row >= 1; --row) {
            if (matchesCondition(ws.get(), row, condition)) {
                rows_to_delete.push_back(row);
            }
        }
        
        // 删除行
        for (int row : rows_to_delete) {
            ws->deleteRows(row, 1);
        }
        
        return rows_to_delete.size();
    }
    
    // 数据统计
    struct TableStatistics {
        size_t record_count;
        std::map<std::string, size_t> value_counts;  // 每列的非空值数量
        std::map<std::string, double> numeric_sums;   // 数值列的总和
    };
    
    TableStatistics getStatistics(const std::string& table_name) {
        auto ws = getWorksheet(table_name);
        TableStatistics stats;
        
        if (!ws) return stats;
        
        auto used_range = ws->getUsedRange();
        stats.record_count = used_range.last_row;
        
        // 获取列头
        auto headers = ws->getRange<std::string>("A1:" + char('A' + used_range.last_col) + "1");
        
        // 统计每列数据
        for (size_t col_idx = 0; col_idx < headers.size(); ++col_idx) {
            const std::string& column_name = headers[col_idx];
            
            std::string col_range = char('A' + col_idx) + "2:" + 
                                   char('A' + col_idx) + std::to_string(used_range.last_row + 1);
            
            // 统计非空值
            auto values = ws->getRange<std::string>(col_range);
            size_t non_empty_count = 0;
            double numeric_sum = 0.0;
            
            for (const auto& value : values) {
                if (!value.empty()) {
                    non_empty_count++;
                    
                    // 尝试转换为数值
                    try {
                        double numeric_value = std::stod(value);
                        numeric_sum += numeric_value;
                    } catch (...) {
                        // 不是数值，忽略
                    }
                }
            }
            
            stats.value_counts[column_name] = non_empty_count;
            if (numeric_sum != 0.0) {
                stats.numeric_sums[column_name] = numeric_sum;
            }
        }
        
        return stats;
    }
    
    void save() {
        workbook_->save();
    }

private:
    std::shared_ptr<Worksheet> getWorksheet(const std::string& name) {
        auto it = worksheets_.find(name);
        return (it != worksheets_.end()) ? it->second : nullptr;
    }
    
    template<typename RecordType>
    void insertRecordAtRow(Worksheet* ws, int row, const RecordType& record) {
        // 使用反射或模板特化来提取记录字段
        // 这里简化为示例代码
        if constexpr (std::is_same_v<RecordType, Product>) {
            ws->set(row, 0, record.code)
              .set(row, 1, record.name)
              .set(row, 2, record.price)
              .set(row, 3, record.stock);
        }
    }
    
    int getHeaderStyleId() {
        // 获取或创建表头样式
        static int header_style_id = -1;
        if (header_style_id == -1) {
            auto style = workbook_->createStyleBuilder()
                .font().bold(true).color(Color::WHITE)
                .fill().pattern(FillPattern::SOLID).color(Color::DARK_BLUE)
                .border().all(BorderStyle::THIN)
                .build();
            header_style_id = workbook_->addStyle(style);
        }
        return header_style_id;
    }
};

// 使用示例
void manageDynamicData() {
    DynamicDataManager manager("dynamic_data.xlsx");
    
    // 定义产品表结构
    std::vector<ColumnDefinition> product_columns = {
        {"产品编码", 15.0, ValidationRule{ValidationType::TEXT_LENGTH, "1", "20"}},
        {"产品名称", 25.0, ValidationRule{ValidationType::TEXT_LENGTH, "1", "50"}},
        {"价格", 10.0, ValidationRule{ValidationType::DECIMAL, "0", "99999.99"}},
        {"库存", 8.0, ValidationRule{ValidationType::WHOLE, "0", "10000"}}
    };
    
    manager.createDataTable("产品表", product_columns);
    
    // 批量插入产品数据
    std::vector<Product> products = loadProducts();
    manager.insertRecords("产品表", products);
    
    // 查询低库存产品
    QueryCondition low_stock_condition;
    low_stock_condition.column = "库存";
    low_stock_condition.operator_type = QueryOperator::LESS_THAN;
    low_stock_condition.value = "10";
    
    auto low_stock_products = manager.queryRecords<Product>("产品表", low_stock_condition);
    
    // 更新价格
    Product price_update;
    price_update.price = 299.99;
    
    QueryCondition iphone_condition;
    iphone_condition.column = "产品名称";
    iphone_condition.operator_type = QueryOperator::CONTAINS;
    iphone_condition.value = "iPhone";
    
    int updated = manager.updateRecords("产品表", iphone_condition, price_update);
    std::cout << "更新了 " << updated << " 个iPhone产品的价格" << std::endl;
    
    // 获取统计信息
    auto stats = manager.getStatistics("产品表");
    std::cout << "产品总数: " << stats.record_count << std::endl;
    std::cout << "总库存价值: " << stats.numeric_sums["价格"] << std::endl;
    
    manager.save();
}
```

## 性能优化建议

### 1. 选择合适的打开模式

```cpp
// 只读取数据 - 使用只读模式
auto wb = Workbook::openForReading("large_data.xlsx");  // 内存占用更少，加载更快

// 需要编辑 - 使用编辑模式
auto wb = Workbook::openForEditing("data.xlsx");        // 完整功能，支持修改

// 创建新文件 - 直接创建
auto wb = Workbook::create("new_file.xlsx");            // 最高性能
```

### 2. 批量操作优于逐个操作

```cpp
// ❌ 低效：逐个操作
for (int i = 0; i < 1000; ++i) {
    ws->setValue(i, 0, data[i]);  // 1000次函数调用
}

// ✅ 高效：批量操作
ws->setRange("A1:A1000", data);  // 1次调用，性能提升5-10倍
```

### 3. 重用工作表引用

```cpp
// ❌ 低效：重复获取工作表
for (int i = 0; i < 100; ++i) {
    workbook->sheet("数据表")->setValue(i, 0, data[i]);  // 重复查找工作表
}

// ✅ 高效：缓存工作表引用
auto ws = workbook->sheet("数据表");  // 获取一次
for (int i = 0; i < 100; ++i) {
    ws->setValue(i, 0, data[i]);      // 重用引用
}
```

### 4. 合理设置工作模式

```cpp
// 大数据批处理
wb->setMode(Workbook::Mode::BATCH);
wb->setMemoryLimit(256 * 1024 * 1024);  // 256MB限制

// 交互式应用
wb->setMode(Workbook::Mode::INTERACTIVE);

// 流式处理超大文件
wb->setMode(Workbook::Mode::STREAMING);
```

## 错误处理最佳实践

### 1. 异常安全的代码

```cpp
try {
    auto wb = Workbook::open("data.xlsx");
    auto ws = wb->sheet("数据表");
    
    // 安全的数据访问
    if (auto value = ws->tryGetValue<double>("B2")) {
        std::cout << "数值: " << *value << std::endl;
    } else {
        std::cout << "B2不是有效的数值" << std::endl;
    }
    
    // 带默认值的访问
    std::string name = ws->getValueOr<std::string>("A1", "未知");
    double price = ws->getValueOr<double>("B1", 0.0);
    
} catch (const FastExcelException& e) {
    std::cerr << "FastExcel错误: " << e.what() << std::endl;
} catch (const std::exception& e) {
    std::cerr << "系统错误: " << e.what() << std::endl;
}
```

### 2. 资源管理

```cpp
// RAII 自动资源管理
{
    auto wb = Workbook::create("temp.xlsx");
    auto ws = wb->addWorksheet("临时数据");
    
    // 进行操作...
    ws->setValue("A1", "测试数据");
    
    // 自动保存和清理
    wb->save();
}  // wb 在此处自动析构，释放资源
```

## 迁移指南

### 从旧API迁移到新API

| 旧API | 新API | 说明 |
|-------|-------|------|
| `cell.getStringValue()` | `cell.getValue<std::string>()` | 类型安全的模板方法 |
| `workbook->getValue<T>(sheet_idx, row, col)` | `workbook->sheet(sheet_idx)->getValue<T>(row, col)` | 职责分离 |
| 逐个设置单元格 | `ws->setRange("A1:C3", values)` | 批量操作提升性能 |
| 复杂的类型判断 | `ws->tryGetValue<T>()` | 安全访问，避免异常 |

### 迁移步骤

1. **评估现有代码** - 识别需要迁移的API调用
2. **更新包含头文件** - 确保包含新的头文件
3. **替换API调用** - 逐步替换为新API
4. **测试验证** - 确保功能正确性
5. **性能优化** - 利用新API的批量操作特性

## 总结

新的 FastExcel API 设计实现了：

1. **清晰的职责分离** - Workbook 管理工作簿，Worksheet 管理单元格
2. **直观的使用体验** - 符合 Excel 用户的思维模型
3. **类型安全保证** - 编译期类型检查，运行期安全访问
4. **卓越的性能** - 批量操作，内存优化，缓存策略
5. **完整的功能覆盖** - 从基础操作到高级特性
6. **优秀的扩展性** - 支持自定义类型，插件机制

这套API设计将显著提升C++开发者处理Excel文件的效率和体验！

---

*完整指南版本*：v2.0  
*最后更新*：2025-01-11  
*作者*：FastExcel 开发团队