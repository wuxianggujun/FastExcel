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

private:
    GenerationContext context_;

public:
    explicit UnifiedXMLGenerator(const GenerationContext& context);

    // 新增：作为编排器的统一生成入口（批量/流式由 IFileWriter 决定）
    bool generateAll(class ::fastexcel::core::IFileWriter& writer);
    bool generateAll(class ::fastexcel::core::IFileWriter& writer,
                     const class ::fastexcel::core::DirtyManager* dirty_manager);
    // 新增：生成指定部件集合（用于上层精简流程的选择性调度）
    bool generateParts(class ::fastexcel::core::IFileWriter& writer,
                       const std::vector<std::string>& parts_to_generate);
    bool generateParts(class ::fastexcel::core::IFileWriter& writer,
                       const std::vector<std::string>& parts_to_generate,
                       const class ::fastexcel::core::DirtyManager* dirty_manager);

    // 删除所有回调与 string 便捷API，统一使用 IFileWriter 接口

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
    // 无内部直接拼写XML的辅助，全部下沉到各部件生成器

private:
    // 作为 orchestrator 的部件注册与调度
    struct Part;
    std::vector<std::unique_ptr<Part>> parts_;
    void registerDefaultParts();
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
