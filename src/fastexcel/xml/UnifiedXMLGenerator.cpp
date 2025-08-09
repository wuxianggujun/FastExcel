#include "UnifiedXMLGenerator.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/Worksheet.hpp"
#include "WorksheetXMLGenerator.hpp"
#include "StyleSerializer.hpp"
#include "DocPropsXMLGenerator.hpp"
#include "fastexcel/utils/Logger.hpp"

namespace fastexcel {
namespace xml {

// ========== 静态工厂方法实现 ==========

std::unique_ptr<UnifiedXMLGenerator> UnifiedXMLGenerator::fromWorkbook(const core::Workbook* workbook) {
    GenerationContext context;
    context.workbook = workbook;
    context.format_repo = &workbook->getStyleRepository();
    context.sst = workbook->getSharedStringTable(); // 获取SharedStringTable
    
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
    writer.startDocument();
    
    writer.startElement("workbook");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    writer.writeAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    
    // 严格按照libxlsxwriter的fileVersion属性
    writer.startElement("fileVersion");
    writer.writeAttribute("appName", "xl");
    writer.writeAttribute("lastEdited", "4");
    writer.writeAttribute("lowestEdited", "4");
    writer.writeAttribute("rupBuild", "4505");
    writer.endElement(); // fileVersion
    
    // 严格按照libxlsxwriter的workbookPr属性  
    writer.startElement("workbookPr");
    writer.writeAttribute("defaultThemeVersion", "124226");
    writer.endElement(); // workbookPr
    
    // 严格按照libxlsxwriter的bookViews属性
    writer.startElement("bookViews");
    writer.startElement("workbookView");
    writer.writeAttribute("xWindow", "240");
    writer.writeAttribute("yWindow", "15");
    writer.writeAttribute("windowWidth", "16095");
    writer.writeAttribute("windowHeight", "9660");
    // 设置activeTab为0，只激活第一个工作表
    auto sheet_names = context_.workbook->getWorksheetNames();
    if (!sheet_names.empty()) {
        writer.writeAttribute("activeTab", "0");
    }
    writer.endElement(); // workbookView
    writer.endElement(); // bookViews
    
    // 生成sheets部分 - 使用经过测试的结构
    generateWorkbookSheetsSection(writer);
    
    // 严格按照libxlsxwriter的calcPr属性
    writer.startElement("calcPr");
    writer.writeAttribute("calcId", "124519");
    writer.writeAttribute("fullCalcOnLoad", "1");
    writer.endElement(); // calcPr
    
    writer.endElement(); // workbook
    writer.flushBuffer();
}

void UnifiedXMLGenerator::generateWorksheetXML(const core::Worksheet* worksheet,
                                              const std::function<void(const char*, size_t)>& callback) {
    if (!worksheet) {
        return;
    }

    // 委托给独立的WorksheetXMLGenerator
    auto generator = std::make_unique<WorksheetXMLGenerator>(worksheet);
    generator->generate(callback);
}

void UnifiedXMLGenerator::generateStylesXML(const std::function<void(const char*, size_t)>& callback) {
    if (!context_.format_repo) {
        return;
    }
    
    // 使用现有的StyleSerializer
    StyleSerializer::serialize(*context_.format_repo, callback);
}

void UnifiedXMLGenerator::generateSharedStringsXML(const std::function<void(const char*, size_t)>& callback) {
    LOG_DEBUG("UnifiedXMLGenerator::generateSharedStringsXML called");
    LOG_DEBUG("context_.sst = {}", context_.sst ? "not null" : "null");
    
    if (!context_.sst) {
        // 生成空的SST
        LOG_DEBUG("Generating empty SharedStrings XML");
        XMLStreamWriter writer(callback);
        writeXMLHeader(writer);
        
        writer.startElement("sst");
        writeExcelNamespaces(writer, "sst");
        writer.writeAttribute("count", 0);
        writer.writeAttribute("uniqueCount", 0);
        writer.endElement(); // sst
        writer.flushBuffer();
        LOG_DEBUG("Empty SharedStrings XML generated");
        return;
    }
    
    // 使用SharedStringTable的现有方法
    LOG_DEBUG("Delegating to SharedStringTable::generateXML");
    context_.sst->generateXML(callback);
    LOG_DEBUG("SharedStringTable::generateXML completed");
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
    writer.startElement("Default");
    writer.writeAttribute("Extension", "rels");
    writer.writeAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
    writer.endElement(); // Default
    
    writer.startElement("Default");
    writer.writeAttribute("Extension", "xml");
    writer.writeAttribute("ContentType", "application/xml");
    writer.endElement(); // Default
    
    // 覆盖类型
    writer.startElement("Override");
    writer.writeAttribute("PartName", "/xl/workbook.xml");
    writer.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
    writer.endElement(); // Override
    
    writer.startElement("Override");
    writer.writeAttribute("PartName", "/xl/styles.xml");
    writer.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
    writer.endElement(); // Override
    
    // 工作表内容类型
    auto sheet_names = context_.workbook->getWorksheetNames();
    for (size_t i = 0; i < sheet_names.size(); ++i) {
        writer.startElement("Override");
        writer.writeAttribute("PartName", "/xl/worksheets/sheet" + std::to_string(i + 1) + ".xml");
        writer.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
        writer.endElement(); // Override
    }
    
    // 共享字符串（如果启用）
    if (context_.workbook->getOptions().use_shared_strings) {
        writer.startElement("Override");
        writer.writeAttribute("PartName", "/xl/sharedStrings.xml");
        writer.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
        writer.endElement(); // Override
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
        writer.startElement("Relationship");
        writer.writeAttribute("Id", "rId1");
        writer.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
        writer.writeAttribute("Target", "xl/workbook.xml");
        writer.endElement(); // Relationship
        
        writer.startElement("Relationship");
        writer.writeAttribute("Id", "rId2");
        writer.writeAttribute("Type", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties");
        writer.writeAttribute("Target", "docProps/core.xml");
        writer.endElement(); // Relationship
        
        writer.startElement("Relationship");
        writer.writeAttribute("Id", "rId3");
        writer.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
        writer.writeAttribute("Target", "docProps/app.xml");
        writer.endElement(); // Relationship
        
    } else if (rel_type == "workbook" && context_.workbook) {
        // 工作簿关系文件
        int rId = 1;
        
        // 工作表关系
        auto sheet_names = context_.workbook->getWorksheetNames();
        for (size_t i = 0; i < sheet_names.size(); ++i) {
            writer.startElement("Relationship");
            writer.writeAttribute("Id", "rId" + std::to_string(rId++));
            writer.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
            writer.writeAttribute("Target", "worksheets/sheet" + std::to_string(i + 1) + ".xml");
            writer.endElement(); // Relationship
        }
        
        // 样式关系
        writer.startElement("Relationship");
        writer.writeAttribute("Id", "rId" + std::to_string(rId++));
        writer.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
        writer.writeAttribute("Target", "styles.xml");
        writer.endElement(); // Relationship
        
        // 共享字符串关系（如果启用）
        if (context_.workbook->getOptions().use_shared_strings) {
            writer.startElement("Relationship");
            writer.writeAttribute("Id", "rId" + std::to_string(rId++));
            writer.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
            writer.writeAttribute("Target", "sharedStrings.xml");
            writer.endElement(); // Relationship
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
        LOG_WARN("UnifiedXMLGenerator::generateDocPropsXML - workbook is null");
        return;
    }
    
    // 使用专门的DocPropsXMLGenerator处理文档属性XML生成
    if (prop_type == "core") {
        DocPropsXMLGenerator::generateCoreXML(context_.workbook, callback);
    } else if (prop_type == "app") {
        DocPropsXMLGenerator::generateAppXML(context_.workbook, callback);
    } else if (prop_type == "custom") {
        DocPropsXMLGenerator::generateCustomXML(context_.workbook, callback);
    } else {
        LOG_ERROR("UnifiedXMLGenerator: Unknown doc props type: {}", prop_type);
    }
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
        writer.writeAttribute("xmlns:mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
        writer.writeAttribute("mc:Ignorable", "x14ac xr xr2 xr3");
        writer.writeAttribute("xmlns:x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
        writer.writeAttribute("xmlns:xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
        writer.writeAttribute("xmlns:xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");
        writer.writeAttribute("xmlns:xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3");
        writer.writeAttribute("xr:uid", "{2DF72439-4A71-4CEC-B11D-A1BD50977D01}");
    } else if (ns_type == "sst") {
        writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    }
}

std::string UnifiedXMLGenerator::escapeXMLText(const std::string& text) {
    std::string result;
    result.reserve(static_cast<size_t>(text.length() * 1.1)); // 预分配稍多空间
    
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
    
    // 使用与Workbook::generateWorkbookXML相同的逻辑
    auto sheet_names = context_.workbook->getWorksheetNames();
    for (size_t i = 0; i < sheet_names.size(); ++i) {
        // 通过名称获取worksheet对象以获取正确的sheetId
        auto worksheet = context_.workbook->getWorksheet(sheet_names[i]);
        
        writer.startElement("sheet");
        writer.writeAttribute("name", sheet_names[i]);
        if (worksheet) {
            writer.writeAttribute("sheetId", std::to_string(worksheet->getSheetId()));
        } else {
            // 备用方案：使用索引+1作为sheetId
            writer.writeAttribute("sheetId", std::to_string(i + 1));
        }
        writer.writeAttribute("r:id", "rId" + std::to_string(i + 1));
        writer.endElement(); // sheet
    }
    
    writer.endElement(); // sheets
}

void UnifiedXMLGenerator::generateWorkbookPropertiesSection(XMLStreamWriter& writer) {
    (void)writer; // 避免未使用参数警告
    // 生成工作簿属性（如果需要）
    // 这里可以根据需要添加更多属性
}

void UnifiedXMLGenerator::generateWorksheetDataSection(XMLStreamWriter& writer, const core::Worksheet* worksheet) {
    if (!worksheet) return;
    
    writer.startElement("sheetData");
    
    // 获取使用范围
    auto [max_row, max_col] = worksheet->getUsedRange();
    if (max_row < 0 || max_col < 0) {
        writer.endElement(); // sheetData (empty)
        return;
    }
    
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
        writer.writeAttribute("r", std::to_string(row + 1).c_str());
        
        for (int col = 0; col <= max_col; ++col) {
            if (worksheet->hasCellAt(row, col)) {
                const auto& cell = worksheet->getCell(row, col);
                
                writer.startElement("c");
                // 单元格引用
                std::string cell_ref = utils::CommonUtils::cellReference(row, col);
                writer.writeAttribute("r", cell_ref.c_str());
                
                // 样式信息
                if (cell.hasFormat()) {
                    int xf_index = -1;
                    auto format_descriptor = cell.getFormatDescriptor();
                    if (format_descriptor && context_.workbook) {
                        auto& format_repo = context_.workbook->getStyleRepository();
                        for (size_t i = 0; i < format_repo.getFormatCount(); ++i) {
                            auto stored_format = format_repo.getFormat(static_cast<int>(i));
                            if (stored_format && *stored_format == *format_descriptor) {
                                xf_index = static_cast<int>(i);
                                break;
                            }
                        }
                    }
                    if (xf_index >= 0) {
                        writer.writeAttribute("s", std::to_string(xf_index).c_str());
                    }
                }
                
                // 单元格值和类型
                if (cell.isFormula()) {
                    writer.writeAttribute("t", "str");
                    writer.startElement("f");
                    writer.writeText(cell.getFormula().c_str());
                    writer.endElement(); // f
                    // 公式结果是数值类型
                    double result = cell.getFormulaResult();
                    if (result != 0.0) {
                        writer.startElement("v");
                        writer.writeText(std::to_string(result).c_str());
                        writer.endElement(); // v
                    }
                } else if (cell.isString()) {
                    // 根据工作簿设置决定使用共享字符串还是内联字符串
                    if (context_.workbook && context_.workbook->getOptions().use_shared_strings && context_.sst) {
                        writer.writeAttribute("t", "s");
                        writer.startElement("v");
                        int sst_index = context_.sst->getStringId(cell.getStringValue());
                        if (sst_index < 0) {
                            // 如果字符串不在SST中，添加它
                            sst_index = const_cast<core::SharedStringTable*>(context_.sst)->addString(cell.getStringValue());
                        }
                        writer.writeText(std::to_string(sst_index).c_str());
                        writer.endElement(); // v
                    } else {
                        writer.writeAttribute("t", "inlineStr");
                        writer.startElement("is");
                        writer.startElement("t");
                        writer.writeText(escapeXMLText(cell.getStringValue()).c_str());
                        writer.endElement(); // t
                        writer.endElement(); // is
                    }
                } else if (cell.isNumber()) {
                    writer.startElement("v");
                    writer.writeText(std::to_string(cell.getNumberValue()).c_str());
                    writer.endElement(); // v
                } else if (cell.isBoolean()) {
                    writer.writeAttribute("t", "b");
                    writer.startElement("v");
                    writer.writeText(cell.getBooleanValue() ? "1" : "0");
                    writer.endElement(); // v
                }
                
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
    writer.writeAttribute("count", std::to_string(merge_ranges.size()).c_str());
    
    for (const auto& range : merge_ranges) {
        writer.startElement("mergeCell");
        // 将合并范围转换为Excel引用格式 (例如：A1:B2)
        std::string range_ref = utils::CommonUtils::cellReference(range.first_row, range.first_col) + ":" +
                               utils::CommonUtils::cellReference(range.last_row, range.last_col);
        writer.writeAttribute("ref", range_ref.c_str());
        writer.endElement(); // mergeCell
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