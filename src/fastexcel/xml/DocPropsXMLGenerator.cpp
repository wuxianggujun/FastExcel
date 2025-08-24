#include "fastexcel/utils/Logger.hpp"
#include "DocPropsXMLGenerator.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/utils/TimeUtils.hpp"
#include "fastexcel/utils/XMLUtils.hpp"
#include <sstream>
#include <iomanip>

namespace fastexcel {
namespace xml {

// 公共静态方法实现

void DocPropsXMLGenerator::generateCoreXML(const core::Workbook* workbook,
                                           const std::function<void(const std::string&)>& callback) {
    if (!workbook) {
        FASTEXCEL_LOG_WARN("DocPropsXMLGenerator::generateCoreXML - workbook is null");
        return;
    }

    XMLStreamWriter writer(callback);
    writeXMLHeader(writer);

    writer.startElement("cp:coreProperties");
    writer.writeAttribute("xmlns:cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
    writer.writeAttribute("xmlns:dc", "http://purl.org/dc/elements/1.1/");
    writer.writeAttribute("xmlns:dcterms", "http://purl.org/dc/terms/");
    writer.writeAttribute("xmlns:dcmitype", "http://purl.org/dc/dcmitype/");
    writer.writeAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");

    const auto& props = workbook->getDocumentProperties();

    // 标题
    if (!props.title.empty()) {
        writer.startElement("dc:title");
        writer.writeText(utils::XMLUtils::escapeXML(props.title));
        writer.endElement(); // dc:title
    }

    // 主题
    if (!props.subject.empty()) {
        writer.startElement("dc:subject");
        writer.writeText(utils::XMLUtils::escapeXML(props.subject));
        writer.endElement(); // dc:subject
    }

    // 作者
    if (!props.author.empty()) {
        writer.startElement("dc:creator");
        writer.writeText(utils::XMLUtils::escapeXML(props.author));
        writer.endElement(); // dc:creator
    }

    // 关键词
    if (!props.keywords.empty()) {
        writer.startElement("cp:keywords");
        writer.writeText(utils::XMLUtils::escapeXML(props.keywords));
        writer.endElement(); // cp:keywords
    }

    // 描述/注释
    if (!props.comments.empty()) {
        writer.startElement("dc:description");
        writer.writeText(utils::XMLUtils::escapeXML(props.comments));
        writer.endElement(); // dc:description
    }

    // 最后修改者
    writer.startElement("cp:lastModifiedBy");
    writer.writeText("FastExcel Library");
    writer.endElement(); // cp:lastModifiedBy

    // 创建时间
    writer.startElement("dcterms:created");
    writer.writeAttribute("xsi:type", "dcterms:W3CDTF");
    writer.writeText(formatTimeISO8601(props.created_time));
    writer.endElement(); // dcterms:created

    // 修改时间
    writer.startElement("dcterms:modified");
    writer.writeAttribute("xsi:type", "dcterms:W3CDTF");
    writer.writeText(formatTimeISO8601(props.modified_time));
    writer.endElement(); // dcterms:modified

    // 类别
    if (!props.category.empty()) {
        writer.startElement("cp:category");
        writer.writeText(utils::XMLUtils::escapeXML(props.category));
        writer.endElement(); // cp:category
    }

    // 状态
    if (!props.status.empty()) {
        writer.startElement("cp:contentStatus");
        writer.writeText(utils::XMLUtils::escapeXML(props.status));
        writer.endElement(); // cp:contentStatus
    }

    writer.endElement(); // cp:coreProperties
    writer.endDocument();
}

void DocPropsXMLGenerator::generateAppXML(const core::Workbook* workbook,
                                         const std::function<void(const std::string&)>& callback) {
    if (!workbook) {
        FASTEXCEL_LOG_WARN("DocPropsXMLGenerator::generateAppXML - workbook is null");
        return;
    }

    XMLStreamWriter writer(callback);
    writeXMLHeader(writer);

    writer.startElement("Properties");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
    writer.writeAttribute("xmlns:vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");

    // 应用程序信息
    writer.startElement("Application");
    writer.writeText("Microsoft Excel");
    writer.endElement(); // Application

    // 文档安全性
    writer.startElement("DocSecurity");
    writer.writeText("0");
    writer.endElement(); // DocSecurity

    // 缩放裁剪
    writer.startElement("ScaleCrop");
    writer.writeText("false");
    writer.endElement(); // ScaleCrop

    // 生成HeadingPairs和TitlesOfParts
    auto worksheet_names = workbook->getSheetNames();
    generateHeadingPairs(writer, worksheet_names.size());
    generateTitlesOfParts(writer, worksheet_names);

    // 公司信息
    writer.startElement("Company");
    const auto& props = workbook->getDocumentProperties();
    writer.writeText(utils::XMLUtils::escapeXML(props.company.empty() ? "FastExcel Library" : props.company));
    writer.endElement(); // Company

    // 链接更新状态
    writer.startElement("LinksUpToDate");
    writer.writeText("false");
    writer.endElement(); // LinksUpToDate

    // 共享文档
    writer.startElement("SharedDoc");
    writer.writeText("false");
    writer.endElement(); // SharedDoc

    // 超链接更改状态
    writer.startElement("HyperlinksChanged");
    writer.writeText("false");
    writer.endElement(); // HyperlinksChanged

    // 应用程序版本
    writer.startElement("AppVersion");
    writer.writeText("16.0300");
    writer.endElement(); // AppVersion

    writer.endElement(); // Properties
    writer.endDocument();
}

void DocPropsXMLGenerator::generateCustomXML(const core::Workbook* workbook,
                                            const std::function<void(const std::string&)>& callback) {
    if (!workbook) {
        FASTEXCEL_LOG_WARN("DocPropsXMLGenerator::generateCustomXML - workbook is null");
        return;
    }

    // 检查是否有自定义属性
    auto custom_props = workbook->getAllCustomProperties();
    FASTEXCEL_LOG_INFO("DocPropsXMLGenerator::generateCustomXML - found {} custom properties", custom_props.size());
    
    if (custom_props.empty()) {
        FASTEXCEL_LOG_INFO("No custom properties found, skipping custom.xml generation");
        return;
    }

    FASTEXCEL_LOG_INFO("DocPropsXMLGenerator::generateCustomXML - starting XML generation");
    XMLStreamWriter writer(callback);
    writeXMLHeader(writer);

    writer.startElement("Properties");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties");
    writer.writeAttribute("xmlns:vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");

    int pid = 2; // 属性ID从2开始
    for (const auto& prop : custom_props) {
        const std::string& name = prop.first;
        const std::string& value = prop.second;
        
        writer.startElement("property");
        writer.writeAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}");
        writer.writeAttribute("pid", std::to_string(pid++));
        writer.writeAttribute("name", utils::XMLUtils::escapeXML(name));

        // 使用lpwstr类型存储字符串值
        writer.startElement("vt:lpwstr");
        writer.writeText(utils::XMLUtils::escapeXML(value));
        writer.endElement(); // vt:lpwstr

        writer.endElement(); // property
    }

    writer.endElement(); // Properties
    writer.endDocument();
    
    FASTEXCEL_LOG_INFO("DocPropsXMLGenerator::generateCustomXML - XML generation completed successfully");
}

// 私有辅助方法实现

void DocPropsXMLGenerator::writeXMLHeader(XMLStreamWriter& writer) {
    writer.startDocument();
}

std::string DocPropsXMLGenerator::formatTimeISO8601(const std::tm& time) {
    // 使用 TimeUtils 进行时间格式化
    return utils::TimeUtils::formatTimeISO8601(time);
}

void DocPropsXMLGenerator::generateHeadingPairs(XMLStreamWriter& writer, size_t worksheet_count) {
    writer.startElement("HeadingPairs");
    writer.startElement("vt:vector");
    writer.writeAttribute("size", "2");
    writer.writeAttribute("baseType", "variant");

    // 第一对：工作表标题
    writer.startElement("vt:variant");
    writer.startElement("vt:lpstr");
    writer.writeText("工作表");
    writer.endElement(); // vt:lpstr
    writer.endElement(); // vt:variant

    // 第二对：工作表数量
    writer.startElement("vt:variant");
    writer.startElement("vt:i4");
    writer.writeText(std::to_string(worksheet_count));
    writer.endElement(); // vt:i4
    writer.endElement(); // vt:variant

    writer.endElement(); // vt:vector
    writer.endElement(); // HeadingPairs
}

void DocPropsXMLGenerator::generateTitlesOfParts(XMLStreamWriter& writer,
                                                 const std::vector<std::string>& worksheet_names) {
    writer.startElement("TitlesOfParts");
    writer.startElement("vt:vector");
    writer.writeAttribute("size", std::to_string(worksheet_names.size()));
    writer.writeAttribute("baseType", "lpstr");

    for (const auto& name : worksheet_names) {
        writer.startElement("vt:lpstr");
        writer.writeText(utils::XMLUtils::escapeXML(name));
        writer.endElement(); // vt:lpstr
    }

    writer.endElement(); // vt:vector
    writer.endElement(); // TitlesOfParts
}

} // namespace xml
} // namespace fastexcel
