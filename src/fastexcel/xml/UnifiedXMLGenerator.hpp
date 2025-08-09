#pragma once

#include "fastexcel/xml/XMLStreamWriter.hpp"
#include "fastexcel/core/FormatRepository.hpp"
#include "fastexcel/core/SharedStringTable.hpp"
#include <functional>
#include <memory>
#include <sstream>
#include <unordered_map>

namespace fastexcel {

// 前向声明
namespace core {
    class Workbook;
    class Worksheet;
}

namespace xml {

/**
 * @brief 统一XML生成器 - 消除重复，集中管理所有XML生成逻辑
 * 
 * 设计原则：
 * 1. 单一职责：只负责XML生成，不涉及业务逻辑
 * 2. 策略模式：支持多种输出方式（callback/string/file）
 * 3. 模板方法：XML生成的通用流程统一处理
 * 4. 工厂方法：根据类型自动选择生成器
 */
class UnifiedXMLGenerator {
public:
    // XML生成上下文，包含所有必要数据
    struct GenerationContext {
        const core::Workbook* workbook = nullptr;
        const core::Worksheet* worksheet = nullptr;
        const core::FormatRepository* format_repo = nullptr;
        const core::SharedStringTable* sst = nullptr;
        std::unordered_map<std::string, std::string> custom_data;
    };

    // 输出策略枚举
    enum class OutputMode {
        CALLBACK,    // 流式回调输出
        STRING,      // 字符串缓冲输出
        DIRECT_FILE  // 直接文件输出
    };

private:
    GenerationContext context_;
    OutputMode output_mode_ = OutputMode::CALLBACK;
    
    // 回调转换器 - 统一实现，避免重复
    static std::string callbackToString(
        const std::function<void(const std::function<void(const char*, size_t)>&)>& generator) {
        std::ostringstream buffer;
        auto callback = [&buffer](const char* data, size_t size) {
            buffer.write(data, static_cast<std::streamsize>(size));
        };
        generator(callback);
        return buffer.str();
    }

public:
    explicit UnifiedXMLGenerator(const GenerationContext& context) 
        : context_(context) {}

    // ========== 主要XML生成方法 ==========
    
    /**
     * @brief 生成Workbook XML
     * @param callback 输出回调函数
     */
    void generateWorkbookXML(const std::function<void(const char*, size_t)>& callback);
    
    /**
     * @brief 生成Worksheet XML
     * @param worksheet 工作表对象
     * @param callback 输出回调函数
     */
    void generateWorksheetXML(const core::Worksheet* worksheet, 
                             const std::function<void(const char*, size_t)>& callback);
    
    /**
     * @brief 生成样式 XML
     * @param callback 输出回调函数
     */
    void generateStylesXML(const std::function<void(const char*, size_t)>& callback);
    
    /**
     * @brief 生成共享字符串 XML
     * @param callback 输出回调函数
     */
    void generateSharedStringsXML(const std::function<void(const char*, size_t)>& callback);
    
    /**
     * @brief 生成内容类型 XML
     * @param callback 输出回调函数
     */
    void generateContentTypesXML(const std::function<void(const char*, size_t)>& callback);
    
    /**
     * @brief 生成关系 XML
     * @param rel_type 关系类型 ("root" 或 "workbook")
     * @param callback 输出回调函数
     */
    void generateRelationshipsXML(const std::string& rel_type,
                                 const std::function<void(const char*, size_t)>& callback);
    
    /**
     * @brief 生成主题 XML
     * @param callback 输出回调函数
     */
    void generateThemeXML(const std::function<void(const char*, size_t)>& callback);
    
    /**
     * @brief 生成文档属性 XML
     * @param prop_type 属性类型 ("core" 或 "app")
     * @param callback 输出回调函数
     */
    void generateDocPropsXML(const std::string& prop_type,
                            const std::function<void(const char*, size_t)>& callback);

    // ========== 便捷的字符串返回方法 ==========
    
    std::string generateWorkbookXMLString() {
        return callbackToString([this](const auto& cb) { generateWorkbookXML(cb); });
    }
    
    std::string generateWorksheetXMLString(const core::Worksheet* worksheet) {
        return callbackToString([this, worksheet](const auto& cb) { 
            generateWorksheetXML(worksheet, cb); 
        });
    }
    
    std::string generateStylesXMLString() {
        return callbackToString([this](const auto& cb) { generateStylesXML(cb); });
    }
    
    std::string generateSharedStringsXMLString() {
        return callbackToString([this](const auto& cb) { generateSharedStringsXML(cb); });
    }
    
    std::string generateContentTypesXMLString() {
        return callbackToString([this](const auto& cb) { generateContentTypesXML(cb); });
    }
    
    std::string generateRelationshipsXMLString(const std::string& rel_type) {
        return callbackToString([this, &rel_type](const auto& cb) { 
            generateRelationshipsXML(rel_type, cb); 
        });
    }

    // ========== 工厂方法 - 简化创建过程 ==========
    
    /**
     * @brief 从Workbook创建生成器
     * @param workbook 工作簿对象
     * @return 生成器实例
     */
    static std::unique_ptr<UnifiedXMLGenerator> fromWorkbook(const core::Workbook* workbook);
    
    /**
     * @brief 从Worksheet创建生成器
     * @param worksheet 工作表对象
     * @return 生成器实例
     */
    static std::unique_ptr<UnifiedXMLGenerator> fromWorksheet(const core::Worksheet* worksheet);

private:
    // ========== 通用XML生成辅助方法 ==========
    
    /**
     * @brief 写入标准XML头部
     * @param writer XML写入器
     */
    void writeXMLHeader(XMLStreamWriter& writer);
    
    /**
     * @brief 写入Excel命名空间声明
     * @param writer XML写入器
     * @param ns_type 命名空间类型
     */
    void writeExcelNamespaces(XMLStreamWriter& writer, const std::string& ns_type);
    
    /**
     * @brief 转义XML文本
     * @param text 原始文本
     * @return 转义后的文本
     */
    std::string escapeXMLText(const std::string& text);
    
    /**
     * @brief 验证生成的XML内容
     * @param xml_content XML内容
     * @return 是否有效
     */
    bool validateXMLContent(const std::string& xml_content);
    
    // ========== 专用生成方法 ==========
    
    void generateWorkbookSheetsSection(XMLStreamWriter& writer);
    void generateWorkbookPropertiesSection(XMLStreamWriter& writer);
    void generateWorksheetDataSection(XMLStreamWriter& writer, const core::Worksheet* worksheet);
    void generateWorksheetColumnsSection(XMLStreamWriter& writer, const core::Worksheet* worksheet);
    void generateWorksheetMergeCellsSection(XMLStreamWriter& writer, const core::Worksheet* worksheet);
};

/**
 * @brief XML生成器工厂 - 提供统一的创建入口
 */
class XMLGeneratorFactory {
public:
    /**
     * @brief 创建统一的XML生成器
     * @param workbook 工作簿对象
     * @return 生成器实例
     */
    static std::unique_ptr<UnifiedXMLGenerator> createGenerator(const core::Workbook* workbook) {
        return UnifiedXMLGenerator::fromWorkbook(workbook);
    }
    
    /**
     * @brief 创建轻量级XML生成器（用于简单场景）
     * @return 轻量级生成器
     */
    static std::unique_ptr<UnifiedXMLGenerator> createLightweightGenerator();
};

}} // namespace fastexcel::xml