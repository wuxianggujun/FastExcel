#include "UnifiedXMLGenerator.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/Worksheet.hpp"
#include "StyleSerializer.hpp"
#include "fastexcel/utils/Logger.hpp"

namespace fastexcel {
namespace xml {

// ========== 静态工厂方法实现 ==========

std::unique_ptr<UnifiedXMLGenerator> UnifiedXMLGenerator::fromWorkbook(const core::Workbook* workbook) {
    GenerationContext context;
    context.workbook = workbook;
    context.format_repo = &workbook->getStyleRepository();
    // context.sst = workbook->getSharedStringTable(); // 如果有的话
    
    return std::make_unique<UnifiedXMLGenerator>(context);
}

std::unique_ptr<UnifiedXMLGenerator> UnifiedXMLGenerator::fromWorksheet(const core::Worksheet* worksheet) {
    GenerationContext context;
    context.worksheet = worksheet;
    if (worksheet) {
        context.workbook = worksheet->getParentWorkbook().get();
        if (context.workbook) {
            context.format_repo = &context.workbook->getStyleRepository();
        }
    }
    
    return std::make_unique<UnifiedXMLGenerator>(context);
}

// ========== 主要XML生成方法实现 ==========

void UnifiedXMLGenerator::generateWorkbookXML(const std::function<void(const char*, size_t)>& callback) {
    if (!context_.workbook) {
        return;
    }

    XMLStreamWriter writer(callback);
    writeXMLHeader(writer);
    
    writer.startElement("workbook");
    writeExcelNamespaces(writer, "workbook");
    
    // 生成workbook属性
    generateWorkbookPropertiesSection(writer);
    
    // 生成sheets部分
    generateWorkbookSheetsSection(writer);
    
    writer.endElement(); // workbook
    writer.flushBuffer();
}

void UnifiedXMLGenerator::generateWorksheetXML(const core::Worksheet* worksheet,
                                              const std::function<void(const char*, size_t)>& callback) {
    if (!worksheet) {
        return;
    }

    XMLStreamWriter writer(callback);
    writeXMLHeader(writer);
    
    writer.startElement("worksheet");
    writeExcelNamespaces(writer, "worksheet");
    
    // 生成列信息
    generateWorksheetColumnsSection(writer, worksheet);
    
    // 生成数据部分
    generateWorksheetDataSection(writer, worksheet);
    
    // 生成合并单元格
    generateWorksheetMergeCellsSection(writer, worksheet);
    
    writer.endElement(); // worksheet
    writer.flushBuffer();
}

void UnifiedXMLGenerator::generateStylesXML(const std::function<void(const char*, size_t)>& callback) {
    if (!context_.format_repo) {
        return;
    }
    
    // 使用现有的StyleSerializer
    StyleSerializer::serialize(*context_.format_repo, callback);
}

void UnifiedXMLGenerator::generateSharedStringsXML(const std::function<void(const char*, size_t)>& callback) {
    if (!context_.sst) {
        // 生成空的SST
        XMLStreamWriter writer(callback);
        writeXMLHeader(writer);
        
        writer.startElement("sst");
        writeExcelNamespaces(writer, "sst");
        writer.writeAttribute("count", 0);
        writer.writeAttribute("uniqueCount", 0);
        writer.endElement(); // sst
        writer.flushBuffer();
        return;
    }
    
    // 使用SharedStringTable的现有方法
    context_.sst->generateXML(callback);
}

void UnifiedXMLGenerator::generateContentTypesXML(const std::function<void(const char*, size_t)>& callback) {
    if (!context_.workbook) {
        return;
    }

    XMLStreamWriter writer(callback);
    writeXMLHeader(writer);
    
    writer.startElement("Types");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types");
    
    // 默认扩展名类型
    writer.writeEmptyElement("Default");
    writer.writeAttribute("Extension", "rels");
    writer.writeAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
    
    writer.writeEmptyElement("Default");
    writer.writeAttribute("Extension", "xml");
    writer.writeAttribute("ContentType", "application/xml");
    
    // 覆盖类型
    writer.writeEmptyElement("Override");
    writer.writeAttribute("PartName", "/xl/workbook.xml");
    writer.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
    
    writer.writeEmptyElement("Override");
    writer.writeAttribute("PartName", "/xl/styles.xml");
    writer.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
    
    // 工作表内容类型
    auto sheet_names = context_.workbook->getWorksheetNames();
    for (size_t i = 0; i < sheet_names.size(); ++i) {
        writer.writeEmptyElement("Override");
        writer.writeAttribute("PartName", "/xl/worksheets/sheet" + std::to_string(i + 1) + ".xml");
        writer.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
    }
    
    // 共享字符串（如果启用）
    if (context_.workbook->getOptions().use_shared_strings) {
        writer.writeEmptyElement("Override");
        writer.writeAttribute("PartName", "/xl/sharedStrings.xml");
        writer.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
    }
    
    writer.endElement(); // Types
    writer.flushBuffer();
}

void UnifiedXMLGenerator::generateRelationshipsXML(const std::string& rel_type,
                                                   const std::function<void(const char*, size_t)>& callback) {
    XMLStreamWriter writer(callback);
    writeXMLHeader(writer);
    
    writer.startElement("Relationships");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
    
    if (rel_type == "root") {
        // 根关系文件
        writer.writeEmptyElement("Relationship");
        writer.writeAttribute("Id", "rId1");
        writer.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
        writer.writeAttribute("Target", "xl/workbook.xml");
        
        writer.writeEmptyElement("Relationship");
        writer.writeAttribute("Id", "rId2");
        writer.writeAttribute("Type", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties");
        writer.writeAttribute("Target", "docProps/core.xml");
        
        writer.writeEmptyElement("Relationship");
        writer.writeAttribute("Id", "rId3");
        writer.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
        writer.writeAttribute("Target", "docProps/app.xml");
        
    } else if (rel_type == "workbook" && context_.workbook) {
        // 工作簿关系文件
        int rId = 1;
        
        // 工作表关系
        auto sheet_names = context_.workbook->getWorksheetNames();
        for (size_t i = 0; i < sheet_names.size(); ++i) {
            writer.writeEmptyElement("Relationship");
            writer.writeAttribute("Id", "rId" + std::to_string(rId++));
            writer.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
            writer.writeAttribute("Target", "worksheets/sheet" + std::to_string(i + 1) + ".xml");
        }
        
        // 样式关系
        writer.writeEmptyElement("Relationship");
        writer.writeAttribute("Id", "rId" + std::to_string(rId++));
        writer.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
        writer.writeAttribute("Target", "styles.xml");
        
        // 共享字符串关系（如果启用）
        if (context_.workbook->getOptions().use_shared_strings) {
            writer.writeEmptyElement("Relationship");
            writer.writeAttribute("Id", "rId" + std::to_string(rId++));
            writer.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
            writer.writeAttribute("Target", "sharedStrings.xml");
        }
    }
    
    writer.endElement(); // Relationships
    writer.flushBuffer();
}

void UnifiedXMLGenerator::generateThemeXML(const std::function<void(const char*, size_t)>& callback) {
    if (!context_.workbook) {
        return;
    }
    
    // 获取主题XML（如果有自定义主题）
    const std::string& theme_xml = context_.workbook->getThemeXML();
    if (!theme_xml.empty()) {
        callback(theme_xml.c_str(), theme_xml.size());
        return;
    }
    
    // 生成默认主题
    XMLStreamWriter writer(callback);
    writeXMLHeader(writer);
    
    writer.startElement("a:theme");
    writer.writeAttribute("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main");
    writer.writeAttribute("name", "Office Theme");
    
    // 简化的默认主题内容
    writer.writeEmptyElement("a:themeElements");
    
    writer.endElement(); // a:theme
    writer.flushBuffer();
}

void UnifiedXMLGenerator::generateDocPropsXML(const std::string& prop_type,
                                             const std::function<void(const char*, size_t)>& callback) {
    if (!context_.workbook) {
        return;
    }
    
    XMLStreamWriter writer(callback);
    writeXMLHeader(writer);
    
    if (prop_type == "core") {
        writer.startElement("cp:coreProperties");
        writer.writeAttribute("xmlns:cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
        writer.writeAttribute("xmlns:dc", "http://purl.org/dc/elements/1.1/");
        writer.writeAttribute("xmlns:dcterms", "http://purl.org/dc/terms/");
        writer.writeAttribute("xmlns:dcmitype", "http://purl.org/dc/dcmitype/");
        writer.writeAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
        
        const auto& props = context_.workbook->getDocumentProperties();
        
        if (!props.title.empty()) {
            writer.startElement("dc:title");
            writer.writeText(props.title);
            writer.endElement();
        }
        
        if (!props.subject.empty()) {
            writer.startElement("dc:subject");
            writer.writeText(props.subject);
            writer.endElement();
        }
        
        if (!props.author.empty()) {
            writer.startElement("dc:creator");
            writer.writeText(props.author);
            writer.endElement();
        }
        
        writer.endElement(); // cp:coreProperties
        
    } else if (prop_type == "app") {
        writer.startElement("Properties");
        writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
        writer.writeAttribute("xmlns:vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
        
        writer.startElement("Application");
        writer.writeText("FastExcel");
        writer.endElement();
        
        writer.startElement("Company");
        const auto& props = context_.workbook->getDocumentProperties();
        writer.writeText(props.company.empty() ? "FastExcel Library" : props.company);
        writer.endElement();
        
        writer.endElement(); // Properties
    }
    
    writer.flushBuffer();
}

// ========== 辅助方法实现 ==========

void UnifiedXMLGenerator::writeXMLHeader(XMLStreamWriter& writer) {
    writer.startDocument();
}

void UnifiedXMLGenerator::writeExcelNamespaces(XMLStreamWriter& writer, const std::string& ns_type) {
    if (ns_type == "workbook") {
        writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        writer.writeAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    } else if (ns_type == "worksheet") {
        writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        writer.writeAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    } else if (ns_type == "sst") {
        writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    }
}

std::string UnifiedXMLGenerator::escapeXMLText(const std::string& text) {
    std::string result;
    result.reserve(text.length() * 1.1); // 预分配稍多空间
    
    for (char c : text) {
        switch (c) {
            case '&': result += "&amp;"; break;
            case '<': result += "&lt;"; break;
            case '>': result += "&gt;"; break;
            case '"': result += "&quot;"; break;
            case '\'': result += "&apos;"; break;
            default: result += c; break;
        }
    }
    
    return result;
}

bool UnifiedXMLGenerator::validateXMLContent(const std::string& xml_content) {
    // 基本验证：检查XML结构
    return !xml_content.empty() && 
           xml_content.find("<?xml") != std::string::npos;
}

void UnifiedXMLGenerator::generateWorkbookSheetsSection(XMLStreamWriter& writer) {
    if (!context_.workbook) return;
    
    writer.startElement("sheets");
    
    auto sheet_names = context_.workbook->getWorksheetNames();
    for (size_t i = 0; i < sheet_names.size(); ++i) {
        writer.writeEmptyElement("sheet");
        writer.writeAttribute("name", sheet_names[i]);
        writer.writeAttribute("sheetId", static_cast<int>(i + 1));
        writer.writeAttribute("r:id", "rId" + std::to_string(i + 1));
    }
    
    writer.endElement(); // sheets
}

void UnifiedXMLGenerator::generateWorkbookPropertiesSection(XMLStreamWriter& writer) {
    // 生成工作簿属性（如果需要）
    // 这里可以根据需要添加更多属性
}

void UnifiedXMLGenerator::generateWorksheetDataSection(XMLStreamWriter& writer, const core::Worksheet* worksheet) {
    if (!worksheet) return;
    
    writer.startElement("sheetData");
    
    // 获取使用范围
    auto [max_row, max_col] = worksheet->getUsedRange();
    
    // 逐行生成数据
    for (int row = 0; row <= max_row; ++row) {
        bool row_has_data = false;
        
        // 检查行是否有数据
        for (int col = 0; col <= max_col; ++col) {
            if (worksheet->hasCellAt(row, col)) {
                row_has_data = true;
                break;
            }
        }
        
        if (!row_has_data) continue;
        
        writer.startElement("row");
        writer.writeAttribute("r", row + 1);
        
        for (int col = 0; col <= max_col; ++col) {
            if (worksheet->hasCellAt(row, col)) {
                const auto& cell = worksheet->getCell(row, col);
                
                writer.startElement("c");
                // 这里需要实现完整的单元格XML生成
                // 包括单元格引用、类型、值等
                writer.endElement(); // c
            }
        }
        
        writer.endElement(); // row
    }
    
    writer.endElement(); // sheetData
}

void UnifiedXMLGenerator::generateWorksheetColumnsSection(XMLStreamWriter& writer, const core::Worksheet* worksheet) {
    if (!worksheet) return;
    
    const auto& col_info = worksheet->getColumnInfo();
    if (col_info.empty()) return;
    
    writer.startElement("cols");
    
    for (const auto& [col_idx, info] : col_info) {
        writer.writeEmptyElement("col");
        writer.writeAttribute("min", col_idx + 1);
        writer.writeAttribute("max", col_idx + 1);
        if (info.width > 0) {
            writer.writeAttribute("width", info.width);
        }
        if (info.hidden) {
            writer.writeAttribute("hidden", "1");
        }
    }
    
    writer.endElement(); // cols
}

void UnifiedXMLGenerator::generateWorksheetMergeCellsSection(XMLStreamWriter& writer, const core::Worksheet* worksheet) {
    if (!worksheet) return;
    
    const auto& merge_ranges = worksheet->getMergeRanges();
    if (merge_ranges.empty()) return;
    
    writer.startElement("mergeCells");
    writer.writeAttribute("count", static_cast<int>(merge_ranges.size()));
    
    for (const auto& range : merge_ranges) {
        writer.writeEmptyElement("mergeCell");
        // 需要将行列索引转换为Excel引用格式
        // 例如：A1:B2
        // writer.writeAttribute("ref", convertToExcelRange(range));
    }
    
    writer.endElement(); // mergeCells
}

// ========== XMLGeneratorFactory 实现 ==========

std::unique_ptr<UnifiedXMLGenerator> XMLGeneratorFactory::createLightweightGenerator() {
    UnifiedXMLGenerator::GenerationContext context;
    // 轻量级生成器不需要完整的context
    return std::make_unique<UnifiedXMLGenerator>(context);
}

}} // namespace fastexcel::xml