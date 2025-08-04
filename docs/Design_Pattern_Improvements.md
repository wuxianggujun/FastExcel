# FastExcel 设计模式改进建议

## 概述

本文档提出了对FastExcel项目的设计模式改进建议，旨在提高代码的可维护性、可扩展性和性能。

## 1. 工厂模式改进

### 当前问题
- 各种对象创建逻辑分散在不同的类中
- 缺乏统一的对象创建接口

### 改进建议

#### 1.1 抽象工厂模式 - Excel组件工厂

```cpp
// src/fastexcel/factory/ExcelComponentFactory.hpp
namespace fastexcel {
namespace factory {

class AbstractExcelFactory {
public:
    virtual ~AbstractExcelFactory() = default;
    virtual std::unique_ptr<core::Cell> createCell() = 0;
    virtual std::unique_ptr<core::Format> createFormat() = 0;
    virtual std::unique_ptr<core::Worksheet> createWorksheet(const std::string& name) = 0;
};

class StandardExcelFactory : public AbstractExcelFactory {
public:
    std::unique_ptr<core::Cell> createCell() override {
        return std::make_unique<core::Cell>();
    }
    
    std::unique_ptr<core::Format> createFormat() override {
        return std::make_unique<core::Format>();
    }
    
    std::unique_ptr<core::Worksheet> createWorksheet(const std::string& name) override {
        return std::make_unique<core::Worksheet>(name, nullptr);
    }
};

class OptimizedExcelFactory : public AbstractExcelFactory {
public:
    std::unique_ptr<core::Cell> createCell() override {
        auto cell = std::make_unique<core::Cell>();
        // 应用优化配置
        return cell;
    }
    
    std::unique_ptr<core::Format> createFormat() override {
        auto format = std::make_unique<core::Format>();
        // 应用优化配置
        return format;
    }
    
    std::unique_ptr<core::Worksheet> createWorksheet(const std::string& name) override {
        auto worksheet = std::make_unique<core::Worksheet>(name, nullptr);
        worksheet->setOptimizeMode(true);
        return worksheet;
    }
};

}} // namespace fastexcel::factory
```

## 2. 策略模式改进

### 当前问题
- XML生成逻辑硬编码在各个类中
- 缺乏灵活的格式化策略

### 改进建议

#### 2.1 XML生成策略

```cpp
// src/fastexcel/strategy/XMLGenerationStrategy.hpp
namespace fastexcel {
namespace strategy {

class XMLGenerationStrategy {
public:
    virtual ~XMLGenerationStrategy() = default;
    virtual std::string generateXML(const core::Cell& cell) = 0;
    virtual std::string generateXML(const core::Format& format) = 0;
    virtual std::string generateXML(const core::Worksheet& worksheet) = 0;
};

class StandardXMLStrategy : public XMLGenerationStrategy {
public:
    std::string generateXML(const core::Cell& cell) override {
        // 标准XML生成逻辑
        return generateStandardCellXML(cell);
    }
    
    std::string generateXML(const core::Format& format) override {
        return generateStandardFormatXML(format);
    }
    
    std::string generateXML(const core::Worksheet& worksheet) override {
        return generateStandardWorksheetXML(worksheet);
    }
};

class CompactXMLStrategy : public XMLGenerationStrategy {
public:
    std::string generateXML(const core::Cell& cell) override {
        // 紧凑XML生成逻辑（去除不必要的空格和属性）
        return generateCompactCellXML(cell);
    }
    
    std::string generateXML(const core::Format& format) override {
        return generateCompactFormatXML(format);
    }
    
    std::string generateXML(const core::Worksheet& worksheet) override {
        return generateCompactWorksheetXML(worksheet);
    }
};

}} // namespace fastexcel::strategy
```

## 3. 观察者模式改进

### 当前问题
- 格式变更通知机制简单
- 缺乏灵活的事件处理

### 改进建议

#### 3.1 格式变更观察者

```cpp
// src/fastexcel/observer/FormatObserver.hpp
namespace fastexcel {
namespace observer {

enum class FormatChangeType {
    FontChanged,
    AlignmentChanged,
    BorderChanged,
    FillChanged,
    ProtectionChanged
};

class FormatChangeEvent {
public:
    FormatChangeType type;
    const core::Format* format;
    std::string property_name;
    
    FormatChangeEvent(FormatChangeType t, const core::Format* f, const std::string& prop)
        : type(t), format(f), property_name(prop) {}
};

class FormatObserver {
public:
    virtual ~FormatObserver() = default;
    virtual void onFormatChanged(const FormatChangeEvent& event) = 0;
};

class FormatSubject {
private:
    std::vector<FormatObserver*> observers_;
    
public:
    void addObserver(FormatObserver* observer) {
        observers_.push_back(observer);
    }
    
    void removeObserver(FormatObserver* observer) {
        observers_.erase(std::remove(observers_.begin(), observers_.end(), observer), 
                        observers_.end());
    }
    
    void notifyObservers(const FormatChangeEvent& event) {
        for (auto* observer : observers_) {
            observer->onFormatChanged(event);
        }
    }
};

}} // namespace fastexcel::observer
```

## 4. 命令模式改进

### 当前问题
- 缺乏撤销/重做功能
- 批量操作不够灵活

### 改进建议

#### 4.1 单元格操作命令

```cpp
// src/fastexcel/command/CellCommand.hpp
namespace fastexcel {
namespace command {

class Command {
public:
    virtual ~Command() = default;
    virtual void execute() = 0;
    virtual void undo() = 0;
    virtual std::string getDescription() const = 0;
};

class SetCellValueCommand : public Command {
private:
    core::Worksheet* worksheet_;
    int row_, col_;
    std::string new_value_;
    std::string old_value_;
    
public:
    SetCellValueCommand(core::Worksheet* ws, int row, int col, const std::string& value)
        : worksheet_(ws), row_(row), col_(col), new_value_(value) {
        // 保存旧值
        old_value_ = ws->getCell(row, col).getStringValue();
    }
    
    void execute() override {
        worksheet_->writeString(row_, col_, new_value_);
    }
    
    void undo() override {
        worksheet_->writeString(row_, col_, old_value_);
    }
    
    std::string getDescription() const override {
        return "Set cell (" + std::to_string(row_) + "," + std::to_string(col_) + ") to '" + new_value_ + "'";
    }
};

class CommandManager {
private:
    std::vector<std::unique_ptr<Command>> history_;
    size_t current_position_ = 0;
    
public:
    void executeCommand(std::unique_ptr<Command> command) {
        // 清除当前位置之后的历史
        history_.erase(history_.begin() + current_position_, history_.end());
        
        command->execute();
        history_.push_back(std::move(command));
        current_position_ = history_.size();
    }
    
    bool canUndo() const {
        return current_position_ > 0;
    }
    
    bool canRedo() const {
        return current_position_ < history_.size();
    }
    
    void undo() {
        if (canUndo()) {
            --current_position_;
            history_[current_position_]->undo();
        }
    }
    
    void redo() {
        if (canRedo()) {
            history_[current_position_]->execute();
            ++current_position_;
        }
    }
};

}} // namespace fastexcel::command
```

## 5. 装饰器模式改进

### 当前问题
- 功能扩展需要修改原有类
- 缺乏灵活的功能组合

### 改进建议

#### 5.1 工作表装饰器

```cpp
// src/fastexcel/decorator/WorksheetDecorator.hpp
namespace fastexcel {
namespace decorator {

class WorksheetInterface {
public:
    virtual ~WorksheetInterface() = default;
    virtual void writeString(int row, int col, const std::string& value) = 0;
    virtual void writeNumber(int row, int col, double value) = 0;
    virtual core::Cell& getCell(int row, int col) = 0;
};

class WorksheetDecorator : public WorksheetInterface {
protected:
    std::unique_ptr<WorksheetInterface> worksheet_;
    
public:
    explicit WorksheetDecorator(std::unique_ptr<WorksheetInterface> worksheet)
        : worksheet_(std::move(worksheet)) {}
    
    void writeString(int row, int col, const std::string& value) override {
        worksheet_->writeString(row, col, value);
    }
    
    void writeNumber(int row, int col, double value) override {
        worksheet_->writeNumber(row, col, value);
    }
    
    core::Cell& getCell(int row, int col) override {
        return worksheet_->getCell(row, col);
    }
};

class LoggingWorksheetDecorator : public WorksheetDecorator {
public:
    explicit LoggingWorksheetDecorator(std::unique_ptr<WorksheetInterface> worksheet)
        : WorksheetDecorator(std::move(worksheet)) {}
    
    void writeString(int row, int col, const std::string& value) override {
        LOG_DEBUG("Writing string '{}' to cell ({}, {})", value, row, col);
        WorksheetDecorator::writeString(row, col, value);
    }
    
    void writeNumber(int row, int col, double value) override {
        LOG_DEBUG("Writing number {} to cell ({}, {})", value, row, col);
        WorksheetDecorator::writeNumber(row, col, value);
    }
};

class ValidationWorksheetDecorator : public WorksheetDecorator {
public:
    explicit ValidationWorksheetDecorator(std::unique_ptr<WorksheetInterface> worksheet)
        : WorksheetDecorator(std::move(worksheet)) {}
    
    void writeString(int row, int col, const std::string& value) override {
        if (!utils::CommonUtils::isValidCellPosition(row, col)) {
            throw std::invalid_argument("Invalid cell position");
        }
        WorksheetDecorator::writeString(row, col, value);
    }
    
    void writeNumber(int row, int col, double value) override {
        if (!utils::CommonUtils::isValidCellPosition(row, col)) {
            throw std::invalid_argument("Invalid cell position");
        }
        WorksheetDecorator::writeNumber(row, col, value);
    }
};

}} // namespace fastexcel::decorator
```

## 6. 建造者模式改进

### 当前问题
- 复杂对象创建过程分散
- 缺乏链式调用支持

### 改进建议

#### 6.1 工作簿建造者

```cpp
// src/fastexcel/builder/WorkbookBuilder.hpp
namespace fastexcel {
namespace builder {

class WorkbookBuilder {
private:
    std::unique_ptr<core::Workbook> workbook_;
    
public:
    explicit WorkbookBuilder(const std::string& filename)
        : workbook_(core::Workbook::create(filename)) {}
    
    WorkbookBuilder& setTitle(const std::string& title) {
        workbook_->setTitle(title);
        return *this;
    }
    
    WorkbookBuilder& setAuthor(const std::string& author) {
        workbook_->setAuthor(author);
        return *this;
    }
    
    WorkbookBuilder& addWorksheet(const std::string& name) {
        workbook_->addWorksheet(name);
        return *this;
    }
    
    WorkbookBuilder& enableHighPerformanceMode() {
        workbook_->setHighPerformanceMode(true);
        return *this;
    }
    
    WorkbookBuilder& setCompressionLevel(int level) {
        workbook_->setCompressionLevel(level);
        return *this;
    }
    
    std::unique_ptr<core::Workbook> build() {
        return std::move(workbook_);
    }
};

}} // namespace fastexcel::builder
```

## 7. 单例模式改进

### 当前问题
- 全局配置管理分散
- 缺乏线程安全的单例实现

### 改进建议

#### 7.1 配置管理单例

```cpp
// src/fastexcel/config/ConfigManager.hpp
namespace fastexcel {
namespace config {

class ConfigManager {
private:
    static std::unique_ptr<ConfigManager> instance_;
    static std::mutex mutex_;
    
    // 配置项
    bool enable_logging_ = true;
    bool enable_optimization_ = true;
    int default_compression_level_ = 1;
    size_t default_buffer_size_ = 4096;
    
    ConfigManager() = default;
    
public:
    static ConfigManager& getInstance() {
        std::lock_guard<std::mutex> lock(mutex_);
        if (!instance_) {
            instance_ = std::unique_ptr<ConfigManager>(new ConfigManager());
        }
        return *instance_;
    }
    
    // 禁用拷贝和赋值
    ConfigManager(const ConfigManager&) = delete;
    ConfigManager& operator=(const ConfigManager&) = delete;
    
    // 配置访问方法
    bool isLoggingEnabled() const { return enable_logging_; }
    void setLoggingEnabled(bool enabled) { enable_logging_ = enabled; }
    
    bool isOptimizationEnabled() const { return enable_optimization_; }
    void setOptimizationEnabled(bool enabled) { enable_optimization_ = enabled; }
    
    int getDefaultCompressionLevel() const { return default_compression_level_; }
    void setDefaultCompressionLevel(int level) { default_compression_level_ = level; }
    
    size_t getDefaultBufferSize() const { return default_buffer_size_; }
    void setDefaultBufferSize(size_t size) { default_buffer_size_ = size; }
};

}} // namespace fastexcel::config
```

## 8. 模板方法模式改进

### 当前问题
- 文件生成流程固化
- 缺乏可扩展的处理流程

### 改进建议

#### 8.1 文件生成模板

```cpp
// src/fastexcel/template/FileGenerationTemplate.hpp
namespace fastexcel {
namespace template_method {

class FileGenerationTemplate {
public:
    // 模板方法
    bool generateFile(const core::Workbook& workbook) {
        if (!validateWorkbook(workbook)) {
            return false;
        }
        
        prepareGeneration();
        
        if (!generateStructure(workbook)) {
            return false;
        }
        
        if (!generateContent(workbook)) {
            return false;
        }
        
        return finalizeGeneration();
    }
    
protected:
    // 钩子方法 - 子类可以重写
    virtual bool validateWorkbook(const core::Workbook& workbook) {
        return !workbook.getWorksheetNames().empty();
    }
    
    virtual void prepareGeneration() {
        // 默认实现为空
    }
    
    virtual bool finalizeGeneration() {
        return true;
    }
    
    // 抽象方法 - 子类必须实现
    virtual bool generateStructure(const core::Workbook& workbook) = 0;
    virtual bool generateContent(const core::Workbook& workbook) = 0;
};

class StandardFileGenerator : public FileGenerationTemplate {
protected:
    bool generateStructure(const core::Workbook& workbook) override {
        // 生成标准Excel文件结构
        return true;
    }
    
    bool generateContent(const core::Workbook& workbook) override {
        // 生成标准Excel内容
        return true;
    }
};

class OptimizedFileGenerator : public FileGenerationTemplate {
protected:
    void prepareGeneration() override {
        // 优化准备工作
    }
    
    bool generateStructure(const core::Workbook& workbook) override {
        // 生成优化的Excel文件结构
        return true;
    }
    
    bool generateContent(const core::Workbook& workbook) override {
        // 生成优化的Excel内容
        return true;
    }
};

}} // namespace fastexcel::template_method
```

## 总结

这些设计模式改进建议将显著提高FastExcel项目的：

1. **可维护性** - 通过清晰的职责分离和模块化设计
2. **可扩展性** - 通过策略模式和装饰器模式支持功能扩展
3. **可测试性** - 通过依赖注入和接口抽象
4. **性能** - 通过优化的工厂模式和模板方法模式
5. **用户体验** - 通过建造者模式和命令模式提供更好的API

建议按优先级逐步实施这些改进，首先从工厂模式和策略模式开始，然后逐步引入其他模式。