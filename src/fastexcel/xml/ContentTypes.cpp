#include "fastexcel/xml/ContentTypes.hpp"

namespace fastexcel {
namespace xml {

void ContentTypes::addDefault(const std::string& extension, const std::string& content_type) {
    default_types_.push_back({extension, content_type});
}

void ContentTypes::addOverride(const std::string& part_name, const std::string& content_type) {
    override_types_.push_back({part_name, content_type});
}

std::string ContentTypes::generate() const {
    XMLStreamWriter writer;
    writer.startDocument();
    writer.startElement("Types");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types");
    
    // 写入默认类型
    for (const auto& def : default_types_) {
        writer.writeEmptyElement("Default");
        writer.writeAttribute("Extension", def.extension.c_str());
        writer.writeAttribute("ContentType", def.content_type.c_str());
    }
    
    // 写入覆盖类型
    for (const auto& override : override_types_) {
        writer.writeEmptyElement("Override");
        writer.writeAttribute("PartName", override.part_name.c_str());
        writer.writeAttribute("ContentType", override.content_type.c_str());
    }
    
    writer.endElement(); // Types
    writer.endDocument();
    
    return writer.toString();
}

void ContentTypes::clear() {
    default_types_.clear();
    override_types_.clear();
}

void ContentTypes::addExcelDefaults() {
    // 添加Excel文件所需的默认内容类型
    addDefault("rels", "application/vnd.openxmlformats-package.relationships+xml");
    addDefault("xml", "application/xml");
    
    // 添加Excel特定的覆盖类型
    addOverride("/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
    addOverride("/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
    addOverride("/xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
    addOverride("/xl/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml");
    addOverride("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml");
    addOverride("/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml");
}

}} // namespace fastexcel::xml