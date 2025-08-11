#include "fastexcel/utils/ModuleLoggers.hpp"
#include "fastexcel/xml/ContentTypes.hpp"

namespace fastexcel {
namespace xml {

void ContentTypes::addDefault(const std::string& extension, const std::string& content_type) {
    default_types_.push_back({extension, content_type});
}

void ContentTypes::addOverride(const std::string& part_name, const std::string& content_type) {
    override_types_.push_back({part_name, content_type});
}

void ContentTypes::generate(const std::function<void(const char*, size_t)>& callback) const {
    XMLStreamWriter writer(callback);
    writer.startDocument();
    writer.startElement("Types");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types");
    
    // 写入默认类型
    for (const auto& def : default_types_) {
        writer.startElement("Default");
        writer.writeAttribute("Extension", def.extension.c_str());
        writer.writeAttribute("ContentType", def.content_type.c_str());
        writer.endElement(); // Default
    }
    
    // 写入覆盖类型
    for (const auto& override : override_types_) {
        writer.startElement("Override");
        writer.writeAttribute("PartName", override.part_name.c_str());
        writer.writeAttribute("ContentType", override.content_type.c_str());
        writer.endElement(); // Override
    }
    
    writer.endElement(); // Types
    writer.endDocument();
}

void ContentTypes::generateToFile(const std::string& filename) const {
    XMLStreamWriter writer(filename);
    writer.startDocument();
    writer.startElement("Types");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types");
    
    // 写入默认类型
    for (const auto& def : default_types_) {
        writer.startElement("Default");
        writer.writeAttribute("Extension", def.extension.c_str());
        writer.writeAttribute("ContentType", def.content_type.c_str());
        writer.endElement(); // Default
    }
    
    // 写入覆盖类型
    for (const auto& override : override_types_) {
        writer.startElement("Override");
        writer.writeAttribute("PartName", override.part_name.c_str());
        writer.writeAttribute("ContentType", override.content_type.c_str());
        writer.endElement(); // Override
    }
    
    writer.endElement(); // Types
    writer.endDocument();
}

void ContentTypes::clear() {
    default_types_.clear();
    override_types_.clear();
}

void ContentTypes::addExcelDefaults() {
    // 添加Excel文件所需的默认内容类型
    addDefault("rels", "application/vnd.openxmlformats-package.relationships+xml");
    addDefault("xml", "application/xml");
    
    // 注意：不在这里添加具体的Override类型，应该由Workbook类根据实际内容动态添加
    // 这样可以避免添加不存在的文件引用
}

}} // namespace fastexcel::xml