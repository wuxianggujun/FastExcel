# FastExcel Workbook统一接口优化方案

## 问题分析

当前流模式和批量模式存在大量重复代码：
- 两种模式生成相同的XML内容
- 只是写入方式不同（批量 vs 流式）
- 维护成本高，容易出现不一致

## 优化方案：统一XML生成器接口

### 1. 抽象写入接口

```cpp
// 统一的文件写入接口
class IFileWriter {
public:
    virtual ~IFileWriter() = default;
    virtual bool writeFile(const std::string& path, const std::string& content) = 0;
    virtual bool openStreamingFile(const std::string& path) = 0;
    virtual bool writeStreamingChunk(const char* data, size_t size) = 0;
    virtual bool closeStreamingFile() = 0;
};

// 批量模式实现
class BatchFileWriter : public IFileWriter {
private:
    std::vector<std::pair<std::string, std::string>> files_;
    archive::FileManager* file_manager_;
    
public:
    bool writeFile(const std::string& path, const std::string& content) override {
        files_.emplace_back(path, content);
        return true;
    }
    
    bool flush() {
        return file_manager_->writeFiles(std::move(files_));
    }
    
    // 流式方法在批量模式中收集到内存
    bool openStreamingFile(const std::string& path) override {
        current_path_ = path;
        current_content_.clear();
        return true;
    }
    
    bool writeStreamingChunk(const char* data, size_t size) override {
        current_content_.append(data, size);
        return true;
    }
    
    bool closeStreamingFile() override {
        return writeFile(current_path_, current_content_);
    }
};

// 流模式实现
class StreamingFileWriter : public IFileWriter {
private:
    archive::FileManager* file_manager_;
    
public:
    bool writeFile(const std::string& path, const std::string& content) override {
        return file_manager_->writeFile(path, content);
    }
    
    bool openStreamingFile(const std::string& path) override {
        return file_manager_->openStreamingFile(path);
    }
    
    bool writeStreamingChunk(const char* data, size_t size) override {
        return file_manager_->writeStreamingChunk(data, size);
    }
    
    bool closeStreamingFile() override {
        return file_manager_->closeStreamingFile();
    }
};
```

### 2. 统一的Excel结构生成器

```cpp
class ExcelStructureGenerator {
private:
    const Workbook* workbook_;
    std::unique_ptr<IFileWriter> writer_;
    
public:
    ExcelStructureGenerator(const Workbook* workbook, std::unique_ptr<IFileWriter> writer)
        : workbook_(workbook), writer_(std::move(writer)) {}
    
    bool generate() {
        // 统一的生成逻辑，不再重复
        if (!generateBasicFiles()) return false;
        if (!generateWorksheets()) return false;
        
        // 批量模式需要最后flush
        if (auto batch_writer = dynamic_cast<BatchFileWriter*>(writer_.get())) {
            return batch_writer->flush();
        }
        
        return true;
    }
    
private:
    bool generateBasicFiles() {
        // 生成基础文件（Content_Types.xml, _rels/.rels等）
        std::string content_types_xml;
        workbook_->generateContentTypesXML([&content_types_xml](const char* data, size_t size) {
            content_types_xml.append(data, size);
        });
        if (!writer_->writeFile("[Content_Types].xml", content_types_xml)) return false;
        
        // ... 其他基础文件
        return true;
    }
    
    bool generateWorksheets() {
        for (size_t i = 0; i < workbook_->getWorksheetCount(); ++i) {
            auto worksheet = workbook_->getWorksheet(i);
            std::string worksheet_path = workbook_->getWorksheetPath(static_cast<int>(i + 1));
            
            // 对于大工作表使用流式写入，小工作表使用批量写入
            if (shouldUseStreamingForWorksheet(worksheet)) {
                if (!generateWorksheetStreaming(worksheet, worksheet_path)) return false;
            } else {
                if (!generateWorksheetBatch(worksheet, worksheet_path)) return false;
            }
        }
        return true;
    }
    
    bool generateWorksheetStreaming(const std::shared_ptr<Worksheet>& worksheet, const std::string& path) {
        if (!writer_->openStreamingFile(path)) return false;
        
        worksheet->generateXML([this](const char* data, size_t size) {
            writer_->writeStreamingChunk(data, size);
        });
        
        return writer_->closeStreamingFile();
    }
    
    bool generateWorksheetBatch(const std::shared_ptr<Worksheet>& worksheet, const std::string& path) {
        std::string worksheet_xml;
        worksheet->generateXML([&worksheet_xml](const char* data, size_t size) {
            worksheet_xml.append(data, size);
        });
        
        return writer_->writeFile(path, worksheet_xml);
    }
    
    bool shouldUseStreamingForWorksheet(const std::shared_ptr<Worksheet>& worksheet) {
        // 智能决策：大工作表用流式，小工作表用批量
        return worksheet->getCellCount() > 10000;
    }
};
```

### 3. 简化的Workbook接口

```cpp
bool Workbook::generateExcelStructure() {
    // 智能选择写入器
    std::unique_ptr<IFileWriter> writer;
    
    switch (options_.mode) {
        case WorkbookMode::BATCH:
            writer = std::make_unique<BatchFileWriter>(file_manager_.get());
            break;
        case WorkbookMode::STREAMING:
            writer = std::make_unique<StreamingFileWriter>(file_manager_.get());
            break;
        case WorkbookMode::AUTO:
            // 根据数据量自动选择
            if (shouldUseStreaming()) {
                writer = std::make_unique<StreamingFileWriter>(file_manager_.get());
            } else {
                writer = std::make_unique<BatchFileWriter>(file_manager_.get());
            }
            break;
    }
    
    // 使用统一的生成器
    ExcelStructureGenerator generator(this, std::move(writer));
    return generator.generate();
}
```

## 优势

1. **消除重复代码**：XML生成逻辑只写一次
2. **统一接口**：流模式和批量模式使用相同的生成流程
3. **灵活性**：可以混合使用（小文件批量，大文件流式）
4. **可维护性**：修改XML格式只需要改一个地方
5. **可扩展性**：容易添加新的写入模式（如压缩模式、加密模式等）

## 实现步骤

1. 创建`IFileWriter`接口和实现类
2. 创建`ExcelStructureGenerator`统一生成器
3. 重构`Workbook::generateExcelStructure()`使用新接口
4. 移除重复的`generateExcelStructureBatch`和`generateExcelStructureStreaming`方法
5. 测试确保功能一致性

这样就能实现"不用把同样的代码写两次"的目标！