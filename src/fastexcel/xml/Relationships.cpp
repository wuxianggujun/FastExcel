#include "fastexcel/xml/Relationships.hpp"
#include <sstream>

namespace fastexcel {
namespace xml {

void Relationships::addRelationship(const std::string& id, const std::string& type, const std::string& target) {
    addRelationship(id, type, target, "Internal");
}

void Relationships::addRelationship(const std::string& id, const std::string& type, const std::string& target, const std::string& target_mode) {
    relationships_.push_back({id, type, target, target_mode});
}

std::string Relationships::addAutoRelationship(const std::string& type, const std::string& target) {
    std::string id = generateId();
    addRelationship(id, type, target, "Internal");
    return id;
}

std::string Relationships::addAutoRelationship(const std::string& type, const std::string& target, const std::string& target_mode) {
    std::string id = generateId();
    addRelationship(id, type, target, target_mode);
    return id;
}

void Relationships::generate(const std::function<void(const std::string&)>& callback) const {
    XMLStreamWriter writer(callback);
    writer.startDocument();
    writer.startElement("Relationships");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
    
    for (const auto& rel : relationships_) {
        writer.startElement("Relationship");
        writer.writeAttribute("Id", rel.id.c_str());
        writer.writeAttribute("Type", rel.type.c_str());
        writer.writeAttribute("Target", rel.target.c_str());
        if (!rel.target_mode.empty() && rel.target_mode != "Internal") {
            writer.writeAttribute("TargetMode", rel.target_mode.c_str());
        }
        writer.endElement(); // Relationship
    }
    
    writer.endElement(); // Relationships
    writer.endDocument();
}

void Relationships::generateToFile(const std::string& filename) const {
    XMLStreamWriter writer(filename);
    writer.startDocument();
    writer.startElement("Relationships");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
    
    for (const auto& rel : relationships_) {
        writer.startElement("Relationship");
        writer.writeAttribute("Id", rel.id.c_str());
        writer.writeAttribute("Type", rel.type.c_str());
        writer.writeAttribute("Target", rel.target.c_str());
        if (!rel.target_mode.empty() && rel.target_mode != "Internal") {
            writer.writeAttribute("TargetMode", rel.target_mode.c_str());
        }
        writer.endElement(); // Relationship
    }
    
    writer.endElement(); // Relationships
    writer.endDocument();
}

void Relationships::clear() {
    relationships_.clear();
}

std::string Relationships::generateId() const {
    std::ostringstream oss;
    oss << "rId" << (relationships_.size() + 1);
    return oss.str();
}

}} // namespace fastexcel::xml