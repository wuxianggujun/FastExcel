#include "fastexcel/utils/Logger.hpp"
#pragma once

#include "fastexcel/xml/XMLStreamWriter.hpp"
#include <functional>
#include <string>

namespace fastexcel {
namespace core {
    class Workbook;
    struct DocumentProperties;
}

namespace xml {

/**
 * @brief 文档属性XML生成器 - 专门处理docProps相关XML
 * 
 * 职责：
 * - 生成 docProps/core.xml（核心文档属性）
 * - 生成 docProps/app.xml（应用程序属性）
 * - 生成 docProps/custom.xml（自定义属性）
 * 
 * 设计原则：
 * - 单一职责：专注于文档属性XML生成
 * - 开放封闭：易于扩展新的属性类型
 * - 符合Excel OOXML规范的标准结构
 */
class DocPropsXMLGenerator {
public:
    /**
     * @brief 生成docProps/core.xml - 核心文档属性
     * 
     * 包含文档的基本元数据：标题、作者、主题、关键词、
     * 创建时间、修改时间、类别、状态等
     * 
     * @param workbook 工作簿对象
     * @param callback XML输出回调函数
     */
    static void generateCoreXML(const core::Workbook* workbook, 
                                const std::function<void(const char*, size_t)>& callback);

    /**
     * @brief 生成docProps/app.xml - 应用程序属性
     * 
     * 包含应用程序相关信息：应用程序名称、版本、公司信息、
     * 工作表列表、页眉对等信息
     * 
     * @param workbook 工作簿对象  
     * @param callback XML输出回调函数
     */
    static void generateAppXML(const core::Workbook* workbook,
                              const std::function<void(const char*, size_t)>& callback);

    /**
     * @brief 生成docProps/custom.xml - 自定义属性
     * 
     * 包含用户定义的自定义属性键值对
     * 
     * @param workbook 工作簿对象
     * @param callback XML输出回调函数
     */
    static void generateCustomXML(const core::Workbook* workbook,
                                 const std::function<void(const char*, size_t)>& callback);

private:
    /**
     * @brief 写入XML文档头部
     */
    static void writeXMLHeader(XMLStreamWriter& writer);

    /**
     * @brief 格式化时间为ISO8601格式
     * @param time 时间结构体
     * @return 格式化后的时间字符串
     */
    static std::string formatTimeISO8601(const std::tm& time);

    /**
     * @brief 生成HeadingPairs元素（用于app.xml）
     * @param writer XML写入器
     * @param worksheet_count 工作表数量
     */
    static void generateHeadingPairs(XMLStreamWriter& writer, size_t worksheet_count);

    /**
     * @brief 生成TitlesOfParts元素（用于app.xml）
     * @param writer XML写入器
     * @param worksheet_names 工作表名称列表
     */
    static void generateTitlesOfParts(XMLStreamWriter& writer, 
                                     const std::vector<std::string>& worksheet_names);
};

} // namespace xml
} // namespace fastexcel
