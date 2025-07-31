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

std::string Relationships::generate() const {
    XMLWriter writer;
    writer.startDocument();
    writer.startElement("Relationships");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
    
    for (const auto& rel : relationships_) {
        writer.writeEmptyElement("Relationship");
        writer.writeAttribute("Id", rel.id);
        writer.writeAttribute("Type", rel.type);
        writer.writeAttribute("Target", rel.target);
        if (!rel.target_mode.empty() && rel.target_mode != "Internal") {
            writer.writeAttribute("TargetMode", rel.target_mode);
        }
    }
    
    writer.endElement(); // Relationships
    writer.endDocument();
    
    return writer.toString();
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