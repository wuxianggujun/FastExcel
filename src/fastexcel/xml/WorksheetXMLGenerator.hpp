#include "fastexcel/utils/ModuleLoggers.hpp"
#pragma once

#include "fastexcel/xml/XMLStreamWriter.hpp"
#include <functional>
#include <memory>
#include <string>

namespace fastexcel {

// 前向声明
namespace core {
    class Worksheet;
    class Workbook;
    class SharedStringTable;
    class FormatRepository;
    class Cell;  // 添加Cell类的前向声明
}

namespace xml {

/**
 * @brief 工作表XML生成器 - 专门负责生成Worksheet相关的XML
 * 
 * 设计原则：
 * 1. 单一职责：只负责Worksheet XML生成
 * 2. 高内聚：所有Worksheet XML生成逻辑集中在此
 * 3. 低耦合：通过接口与其他组件交互
 * 4. 性能优化：支持批量和流式两种生成模式
 */
class WorksheetXMLGenerator {
public:
    /**
     * @brief 生成模式
     */
    enum class GenerationMode {
        BATCH,      // 批量模式：使用XMLStreamWriter
        STREAMING   // 流式模式：直接字符串拼接
    };

private:
    const core::Worksheet* worksheet_;
    const core::Workbook* workbook_;
    const core::SharedStringTable* sst_;
    const core::FormatRepository* format_repo_;
    GenerationMode mode_;

public:
    /**
     * @brief 构造函数
     * @param worksheet 工作表对象
     */
    explicit WorksheetXMLGenerator(const core::Worksheet* worksheet);
    
    /**
     * @brief 析构函数
     */
    ~WorksheetXMLGenerator() = default;
    
    // 主要生成方法
    
    /**
     * @brief 生成完整的工作表XML
     * @param callback 输出回调函数
     */
    void generate(const std::function<void(const char*, size_t)>& callback);
    
    /**
     * @brief 生成工作表关系XML
     * @param callback 输出回调函数
     */
    void generateRelationships(const std::function<void(const char*, size_t)>& callback);
    
    // 配置方法
    
    /**
     * @brief 设置生成模式
     * @param mode 生成模式
     */
    void setMode(GenerationMode mode) { mode_ = mode; }
    
    /**
     * @brief 获取生成模式
     * @return 当前生成模式
     */
    GenerationMode getMode() const { return mode_; }

private:
    // 批量模式生成方法
    
    /**
     * @brief 批量模式生成工作表XML
     * @param callback 输出回调函数
     */
    void generateBatch(const std::function<void(const char*, size_t)>& callback);
    
    /**
     * @brief 生成工作表视图
     * @param writer XML写入器
     */
    void generateSheetViews(XMLStreamWriter& writer);
    
    /**
     * @brief 生成列信息
     * @param writer XML写入器
     */
    void generateColumns(XMLStreamWriter& writer);
    
    /**
     * @brief 生成单元格数据
     * @param writer XML写入器
     */
    void generateSheetData(XMLStreamWriter& writer);
    
    /**
     * @brief 生成合并单元格
     * @param writer XML写入器
     */
    void generateMergeCells(XMLStreamWriter& writer);
    
    /**
     * @brief 生成自动筛选
     * @param writer XML写入器
     */
    void generateAutoFilter(XMLStreamWriter& writer);
    
    /**
     * @brief 生成工作表保护
     * @param writer XML写入器
     */
    void generateSheetProtection(XMLStreamWriter& writer);
    
    /**
     * @brief 生成打印选项
     * @param writer XML写入器
     */
    void generatePrintOptions(XMLStreamWriter& writer);
    
    /**
     * @brief 生成页面设置
     * @param writer XML写入器
     */
    void generatePageSetup(XMLStreamWriter& writer);
    
    /**
     * @brief 生成页边距
     * @param writer XML写入器
     */
    void generatePageMargins(XMLStreamWriter& writer);
    
    /**
     * @brief 生成图片绘图引用
     * @param writer XML写入器
     */
    void generateDrawing(XMLStreamWriter& writer);
    
    // 流式模式生成方法
    
    /**
     * @brief 流式模式生成工作表XML
     * @param callback 输出回调函数
     */
    void generateStreaming(const std::function<void(const char*, size_t)>& callback);
    
    /**
     * @brief 流式生成单元格数据
     * @param writer XML写入器
     */
    void generateSheetDataStreaming(XMLStreamWriter& writer);
    
    // 辅助方法
    
    /**
     * @brief 获取单元格格式索引
     * @param cell 单元格对象
     * @return 格式索引，-1表示无格式
     */
    int getCellFormatIndex(const core::Cell& cell);
    
    /**
     * @brief 生成单元格XML字符串
     * @param row 行号
     * @param col 列号
     * @param cell 单元格对象
     * @return XML字符串
     */
    /**
     * @brief 生成单元格XML（流式版本，性能更好）
     * @param writer XML写入器
     * @param row 行号
     * @param col 列号
     * @param cell 单元格对象
     */
    void generateCellXMLStreaming(XMLStreamWriter& writer, int row, int col, const core::Cell& cell);
};

/**
 * @brief 工作表XML生成器工厂
 */
class WorksheetXMLGeneratorFactory {
public:
    /**
     * @brief 创建工作表XML生成器
     * @param worksheet 工作表对象
     * @return 生成器实例
     */
    static std::unique_ptr<WorksheetXMLGenerator> create(const core::Worksheet* worksheet) {
        return std::make_unique<WorksheetXMLGenerator>(worksheet);
    }
    
    /**
     * @brief 创建批量模式生成器
     * @param worksheet 工作表对象
     * @return 生成器实例
     */
    static std::unique_ptr<WorksheetXMLGenerator> createBatch(const core::Worksheet* worksheet) {
        auto generator = std::make_unique<WorksheetXMLGenerator>(worksheet);
        generator->setMode(WorksheetXMLGenerator::GenerationMode::BATCH);
        return generator;
    }
    
    /**
     * @brief 创建流式模式生成器
     * @param worksheet 工作表对象
     * @return 生成器实例
     */
    static std::unique_ptr<WorksheetXMLGenerator> createStreaming(const core::Worksheet* worksheet) {
        auto generator = std::make_unique<WorksheetXMLGenerator>(worksheet);
        generator->setMode(WorksheetXMLGenerator::GenerationMode::STREAMING);
        return generator;
    }
};

}} // namespace fastexcel::xml
