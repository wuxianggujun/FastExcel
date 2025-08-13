#include "fastexcel/utils/ModuleLoggers.hpp"
#include "IXMLPartGenerator.hpp"
#include "UnifiedXMLGenerator.hpp"
#include "WorksheetXMLGenerator.hpp"
#include "DrawingXMLGenerator.hpp"
#include "StyleSerializer.hpp"
#include "DocPropsXMLGenerator.hpp"
#include "fastexcel/core/IFileWriter.hpp"
#include "fastexcel/core/DirtyManager.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/Image.hpp"
#include "fastexcel/theme/Theme.hpp"
#include "fastexcel/xml/XMLStreamWriter.hpp"
#include "fastexcel/utils/Logger.hpp"

using fastexcel::core::IFileWriter;

namespace fastexcel { namespace xml {

static bool writeWithCallback(IFileWriter& writer,
                              const std::string& path,
                              const std::function<void(const std::function<void(const char*, size_t)>&)>& gen) {
    if (!writer.openStreamingFile(path)) {
        XML_ERROR("Failed to open streaming file: {}", path);
        return false;
    }
    
    try {
        auto cb = [&writer](const char* data, size_t sz){ writer.writeStreamingChunk(data, sz); };
        gen(cb);
        return writer.closeStreamingFile();
    } catch (const std::exception& e) {
        XML_ERROR("Exception during streaming generation for {}: {}", path, e.what());
        writer.closeStreamingFile(); // 确保清理状态
        return false;
    }
}

// ContentTypes
class ContentTypesGenerator : public IXMLPartGenerator {
public:
    std::vector<std::string> partNames(const XMLContextView&) const override {
        return {"[Content_Types].xml"};
    }
    bool generatePart(const std::string& part, const XMLContextView& ctx, IFileWriter& writer) override {
        if (part != "[Content_Types].xml") return false;
        return writeWithCallback(writer, part, [&ctx](auto& cb){
            XMLStreamWriter w(cb);
            w.startDocument();
            w.startElement("Types");
            w.writeAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types");
            w.startElement("Default"); w.writeAttribute("Extension", "rels"); w.writeAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml"); w.endElement();
            w.startElement("Default"); w.writeAttribute("Extension", "xml"); w.writeAttribute("ContentType", "application/xml"); w.endElement();
            // 添加图片文件的默认内容类型
            w.startElement("Default"); w.writeAttribute("Extension", "png"); w.writeAttribute("ContentType", "image/png"); w.endElement();
            w.startElement("Default"); w.writeAttribute("Extension", "jpg"); w.writeAttribute("ContentType", "image/jpeg"); w.endElement();
            w.startElement("Default"); w.writeAttribute("Extension", "jpeg"); w.writeAttribute("ContentType", "image/jpeg"); w.endElement();
            // 🔧 关键修复：添加 docProps 的内容类型声明
            w.startElement("Override"); w.writeAttribute("PartName", "/docProps/core.xml"); w.writeAttribute("ContentType", "application/vnd.openxmlformats-package.core-properties+xml"); w.endElement();
            w.startElement("Override"); w.writeAttribute("PartName", "/docProps/app.xml"); w.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.extended-properties+xml"); w.endElement();
            w.startElement("Override"); w.writeAttribute("PartName", "/xl/workbook.xml"); w.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"); w.endElement();
            w.startElement("Override"); w.writeAttribute("PartName", "/xl/styles.xml"); w.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"); w.endElement();
            if (ctx.workbook) {
                auto sheet_names = ctx.workbook->getSheetNames();
                for (size_t i = 0; i < sheet_names.size(); ++i) {
                    w.startElement("Override");
                    w.writeAttribute("PartName", (std::string("/xl/worksheets/sheet") + std::to_string(i+1) + ".xml").c_str());
                    w.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
                    w.endElement();
                    
                    // 添加绘图内容类型（如果工作表包含图片）
                    auto ws = ctx.workbook->getSheet(static_cast<size_t>(i));
                    if (ws && !ws->getImages().empty()) {
                        w.startElement("Override");
                        w.writeAttribute("PartName", (std::string("/xl/drawings/drawing") + std::to_string(i+1) + ".xml").c_str());
                        w.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml");
                        w.endElement();
                    }
                }
                if (ctx.workbook->getOptions().use_shared_strings) {
                    w.startElement("Override");
                    w.writeAttribute("PartName", "/xl/sharedStrings.xml");
                    w.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
                    w.endElement();
                }
            }
            w.endElement();
            w.flushBuffer();
        });
    }
};

// Root relationships
class RootRelsGenerator : public IXMLPartGenerator {
public:
    std::vector<std::string> partNames(const XMLContextView&) const override {
        return {"_rels/.rels"};
    }
    bool generatePart(const std::string& part, const XMLContextView& ctx, IFileWriter& writer) override {
        (void)ctx;
        if (part != "_rels/.rels") return false;
        return writeWithCallback(writer, part, [](auto& cb){
            XMLStreamWriter w(cb); w.startDocument();
            w.startElement("Relationships"); w.writeAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
            auto rel = [&](const char* id, const char* type, const char* target){ w.startElement("Relationship"); w.writeAttribute("Id", id); w.writeAttribute("Type", type); w.writeAttribute("Target", target); w.endElement(); };
            rel("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "xl/workbook.xml");
            rel("rId2", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml");
            rel("rId3", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "docProps/app.xml");
            w.endElement(); w.flushBuffer();
        });
    }
};

// Workbook xml and rels
class WorkbookPartGenerator : public IXMLPartGenerator {
public:
    std::vector<std::string> partNames(const XMLContextView&) const override {
        return {"xl/workbook.xml", "xl/_rels/workbook.xml.rels"};
    }
    bool generatePart(const std::string& part, const XMLContextView& ctx, IFileWriter& writer) override {
        if (part == "xl/_rels/workbook.xml.rels") {
            return writeWithCallback(writer, part, [&ctx](auto& cb){
                XMLStreamWriter w(cb); w.startDocument();
                w.startElement("Relationships"); w.writeAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
                int rId = 1;
                if (ctx.workbook) {
                    auto sheet_names = ctx.workbook->getSheetNames();
                    for (size_t i = 0; i < sheet_names.size(); ++i) {
                        w.startElement("Relationship");
                        w.writeAttribute("Id", (std::string("rId") + std::to_string(rId++)).c_str());
                        w.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
                        w.writeAttribute("Target", (std::string("worksheets/sheet") + std::to_string(i+1) + ".xml").c_str());
                        w.endElement();
                    }
                }
                w.startElement("Relationship"); w.writeAttribute("Id", (std::string("rId") + std::to_string(rId++)).c_str()); w.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"); w.writeAttribute("Target", "styles.xml"); w.endElement();
                if (ctx.workbook && ctx.workbook->getOptions().use_shared_strings) {
                    w.startElement("Relationship"); w.writeAttribute("Id", (std::string("rId") + std::to_string(rId++)).c_str()); w.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"); w.writeAttribute("Target", "sharedStrings.xml"); w.endElement();
                }
                w.endElement(); w.flushBuffer();
            });
        }
        if (part == "xl/workbook.xml") {
            return writeWithCallback(writer, part, [&ctx](auto& cb){
                XMLStreamWriter w(cb); w.startDocument();
                w.startElement("workbook");
                w.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                w.writeAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                w.startElement("workbookPr"); w.writeAttribute("defaultThemeVersion", "124226"); w.endElement();
                w.startElement("bookViews"); w.startElement("workbookView"); w.writeAttribute("xWindow", "240"); w.writeAttribute("yWindow", "15"); w.writeAttribute("windowWidth", "16095"); w.writeAttribute("windowHeight", "9660"); w.writeAttribute("activeTab", "0"); w.endElement(); w.endElement();
                w.startElement("sheets");
                if (ctx.workbook) {
                    auto sheet_names = ctx.workbook->getSheetNames();
                    for (size_t i = 0; i < sheet_names.size(); ++i) {
                        auto ws = ctx.workbook->getSheet(sheet_names[i]);
                        w.startElement("sheet");
                        w.writeAttribute("name", sheet_names[i].c_str());
                        int sheetId = ws ? ws->getSheetId() : static_cast<int>(i+1);
                        w.writeAttribute("sheetId", std::to_string(sheetId).c_str());
                        w.writeAttribute("r:id", (std::string("rId") + std::to_string(i+1)).c_str());
                        w.endElement();
                    }
                }
                w.endElement();
                w.startElement("calcPr"); w.writeAttribute("calcId", "124519"); w.writeAttribute("fullCalcOnLoad", "1"); w.endElement();
                w.endElement(); w.flushBuffer();
            });
        }
        return false;
    }
};

// Styles
class StylesGenerator : public IXMLPartGenerator {
public:
    std::vector<std::string> partNames(const XMLContextView&) const override {
        return {"xl/styles.xml"};
    }
    bool generatePart(const std::string& part, const XMLContextView& ctx, IFileWriter& writer) override {
        if (part != "xl/styles.xml") return false;
        if (!ctx.format_repo) return true; // 无样式也可视为成功（生成空或默认）
        return writeWithCallback(writer, part, [&ctx](auto& cb){ StyleSerializer::serialize(*ctx.format_repo, cb); });
    }
};

// SharedStrings
class SharedStringsGenerator : public IXMLPartGenerator {
public:
    std::vector<std::string> partNames(const XMLContextView& ctx) const override {
        if (!ctx.workbook || !ctx.workbook->getOptions().use_shared_strings) return {};
        return {"xl/sharedStrings.xml"};
    }
    bool generatePart(const std::string& part, const XMLContextView& ctx, IFileWriter& writer) override {
        if (part != "xl/sharedStrings.xml") return false;
        return writeWithCallback(writer, part, [&ctx](auto& cb){
            if (ctx.sst) {
                ctx.sst->generateXML(cb);
            } else {
                XMLStreamWriter w(cb); w.startDocument();
                w.startElement("sst"); w.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"); w.writeAttribute("count", 0); w.writeAttribute("uniqueCount", 0); w.endElement(); w.flushBuffer();
            }
        });
    }
};

// Theme
class ThemeGenerator : public IXMLPartGenerator {
public:
    std::vector<std::string> partNames(const XMLContextView& ctx) const override {
        if (!ctx.theme) return {};
        return {"xl/theme/theme1.xml"};
    }
    bool generatePart(const std::string& part, const XMLContextView& ctx, IFileWriter& writer) override {
        if (part != "xl/theme/theme1.xml") return false;
        if (!ctx.theme) return true; // 无主题则跳过
        std::string xml = ctx.theme->toXML();
        return writer.writeFile(part, xml);
    }
};

// DocProps core/app
class DocPropsGenerator : public IXMLPartGenerator {
public:
    std::vector<std::string> partNames(const XMLContextView&) const override {
        return {"docProps/core.xml", "docProps/app.xml", "docProps/custom.xml"};
    }
    bool generatePart(const std::string& part, const XMLContextView& ctx, IFileWriter& writer) override {
        if (part == "docProps/core.xml") {
            return writeWithCallback(writer, part, [&](auto& cb){ DocPropsXMLGenerator::generateCoreXML(ctx.workbook, cb); });
        }
        if (part == "docProps/app.xml") {
            return writeWithCallback(writer, part, [&](auto& cb){ DocPropsXMLGenerator::generateAppXML(ctx.workbook, cb); });
        }
        if (part == "docProps/custom.xml") {
            return writeWithCallback(writer, part, [&](auto& cb){ DocPropsXMLGenerator::generateCustomXML(ctx.workbook, cb); });
        }
        return false;
    }
};

// Worksheets
class WorksheetsGenerator : public IXMLPartGenerator {
public:
    std::vector<std::string> partNames(const XMLContextView& ctx) const override {
        std::vector<std::string> parts;
        if (!ctx.workbook) return parts;
        auto names = ctx.workbook->getSheetNames();
        for (size_t i = 0; i < names.size(); ++i) {
            parts.emplace_back("xl/worksheets/sheet" + std::to_string(i + 1) + ".xml");
        }
        return parts;
    }
    bool generatePart(const std::string& part, const XMLContextView& ctx, IFileWriter& writer) override {
        if (!ctx.workbook) return false;
        // 解析 sheet index
        // part format: xl/worksheets/sheet{N}.xml
        auto pos1 = part.rfind("sheet");  // 🔧 使用 rfind 找最后一个 "sheet"
        auto pos2 = part.find(".xml");
        if (pos1 == std::string::npos || pos2 == std::string::npos) return false;
        
        // 🔧 关键修复：正确计算数字部分的起始位置
        size_t number_start = pos1 + 5; // "sheet" 有5个字符
        if (number_start >= pos2) return false; // 确保有数字部分
        
        std::string number_str = part.substr(number_start, pos2 - number_start);
        if (number_str.empty()) return false;
        
        int idx;
        try {
            idx = std::stoi(number_str) - 1;
        } catch (const std::exception&) {
            XML_ERROR("Failed to parse sheet index from path: {}, extracted: '{}'", part, number_str);
            return false;
        }
        auto names = ctx.workbook->getSheetNames();
        if (idx < 0 || static_cast<size_t>(idx) >= names.size()) return false;
        auto ws = ctx.workbook->getSheet(static_cast<size_t>(idx));
        if (!ws) return false;

        // 使用现有 WorksheetXMLGenerator 流式输出
        WorksheetXMLGenerator gen(ws.get());
        return writeWithCallback(writer, part, [&](auto& cb){ gen.generate(cb); });
    }
};

// Worksheet rels
class WorksheetRelsGenerator : public IXMLPartGenerator {
public:
    std::vector<std::string> partNames(const XMLContextView& ctx) const override {
        std::vector<std::string> parts;
        if (!ctx.workbook) return parts;
        auto names = ctx.workbook->getSheetNames();
        for (size_t i = 0; i < names.size(); ++i) {
            parts.emplace_back("xl/worksheets/_rels/sheet" + std::to_string(i + 1) + ".xml.rels");
        }
        return parts;
    }
    bool generatePart(const std::string& part, const XMLContextView& ctx, IFileWriter& writer) override {
        if (!ctx.workbook) return false;
        auto pos1 = part.rfind("sheet");  // 🔧 使用 rfind 找最后一个 "sheet"
        auto pos2 = part.find(".xml.rels");
        if (pos1 == std::string::npos || pos2 == std::string::npos) return false;
        
        // 🔧 关键修复：正确计算数字部分的起始位置
        size_t number_start = pos1 + 5; // "sheet" 有5个字符
        if (number_start >= pos2) return false; // 确保有数字部分
        
        std::string number_str = part.substr(number_start, pos2 - number_start);
        if (number_str.empty()) return false;
        
        int idx;
        try {
            idx = std::stoi(number_str) - 1;
        } catch (const std::exception&) {
            XML_ERROR("Failed to parse sheet index from rels path: {}, extracted: '{}'", part, number_str);
            return false;
        }
        auto ws = ctx.workbook->getSheet(static_cast<size_t>(idx));
        if (!ws) return true; // nothing to do
        std::string rels_xml;
        ws->generateRelsXML([&rels_xml](const char* data, size_t size){ rels_xml.append(data, size); });
        if (rels_xml.empty()) return true; // 无关系则不写文件
        return writer.writeFile(part, rels_xml);
    }
};

// Drawing XML Generator
class DrawingPartGenerator : public IXMLPartGenerator {
public:
    std::vector<std::string> partNames(const XMLContextView& ctx) const override {
        std::vector<std::string> parts;
        if (!ctx.workbook) return parts;
        
        // 检查所有工作表是否包含图片
        auto names = ctx.workbook->getSheetNames();
        for (size_t i = 0; i < names.size(); ++i) {
            auto ws = ctx.workbook->getSheet(static_cast<size_t>(i));
            if (ws && !ws->getImages().empty()) {
                parts.emplace_back("xl/drawings/drawing" + std::to_string(i + 1) + ".xml");
            }
        }
        return parts;
    }
    
    bool generatePart(const std::string& part, const XMLContextView& ctx, IFileWriter& writer) override {
        if (!ctx.workbook) return false;
        
        // 解析绘图索引 (xl/drawings/drawing{N}.xml)
        auto pos1 = part.rfind("drawing");
        auto pos2 = part.find(".xml");
        if (pos1 == std::string::npos || pos2 == std::string::npos) return false;
        
        size_t number_start = pos1 + 7; // "drawing" 有7个字符
        if (number_start >= pos2) return false;
        
        std::string number_str = part.substr(number_start, pos2 - number_start);
        if (number_str.empty()) return false;
        
        int idx;
        try {
            idx = std::stoi(number_str) - 1;
        } catch (const std::exception&) {
            XML_ERROR("Failed to parse drawing index from path: {}, extracted: '{}'", part, number_str);
            return false;
        }
        
        auto ws = ctx.workbook->getSheet(static_cast<size_t>(idx));
        if (!ws || ws->getImages().empty()) {
            XML_ERROR("No worksheet or no images for drawing index {}", idx);
            return false;
        }
        
        // 🔧 关键修复：直接生成XML，不检查hasImages（因为已经检查过了）
        const auto& images = ws->getImages();
        XML_DEBUG("Generating drawing XML for {} images", images.size());
        
        // 直接写入XML内容，避免DrawingXMLGenerator的hasImages检查
        return writeWithCallback(writer, part, [&images, idx](auto& cb){
            XMLStreamWriter w(cb);
            w.startDocument();
            
            // 根元素
            w.startElement("xdr:wsDr");
            w.writeAttribute("xmlns:xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            w.writeAttribute("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            w.writeAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            
            // 生成每个图片的XML
            int image_index = 0;
            for (const auto& image : images) {
                if (image) {
                    const auto& anchor = image->getAnchor();
                    
                    // 使用twoCellAnchor以固定图片位置
                    w.startElement("xdr:twoCellAnchor");
                    w.writeAttribute("editAs", "oneCell"); // 固定图片，不允许移动和调整大小
                    
                    // 起始位置
                    w.startElement("xdr:from");
                    w.startElement("xdr:col");
                    w.writeText(std::to_string(anchor.from_col));
                    w.endElement();
                    w.startElement("xdr:colOff");
                    w.writeText("0");
                    w.endElement();
                    w.startElement("xdr:row");
                    w.writeText(std::to_string(anchor.from_row));
                    w.endElement();
                    w.startElement("xdr:rowOff");
                    w.writeText("0");
                    w.endElement();
                    w.endElement(); // xdr:from
                    
                    // 结束位置（根据图片大小计算）
                    // 默认每列宽度64像素，每行高度20像素
                    int to_col = anchor.from_col + static_cast<int>(std::max(100.0, anchor.width) / 64.0) + 1;
                    int to_row = anchor.from_row + static_cast<int>(std::max(100.0, anchor.height) / 20.0) + 1;
                    
                    w.startElement("xdr:to");
                    w.startElement("xdr:col");
                    w.writeText(std::to_string(to_col));
                    w.endElement();
                    w.startElement("xdr:colOff");
                    w.writeText("0");
                    w.endElement();
                    w.startElement("xdr:row");
                    w.writeText(std::to_string(to_row));
                    w.endElement();
                    w.startElement("xdr:rowOff");
                    w.writeText("0");
                    w.endElement();
                    w.endElement(); // xdr:to
                    
                    // 图片
                    w.startElement("xdr:pic");
                    
                    // 非可视属性
                    w.startElement("xdr:nvPicPr");
                    w.startElement("xdr:cNvPr");
                    w.writeAttribute("id", std::to_string(image_index + 2));
                    w.writeAttribute("name", image->getName().empty() ? ("Picture " + std::to_string(image_index + 1)) : image->getName());
                    w.endElement(); // xdr:cNvPr
                    w.startElement("xdr:cNvPicPr");
                    w.startElement("a:picLocks");
                    w.writeAttribute("noChangeAspect", "1");
                    w.endElement();
                    w.endElement(); // xdr:cNvPicPr
                    w.endElement(); // xdr:nvPicPr
                    
                    // 图片填充
                    w.startElement("xdr:blipFill");
                    w.startElement("a:blip");
                    w.writeAttribute("r:embed", "rId" + std::to_string(image_index + 1));
                    w.endElement(); // a:blip
                    w.startElement("a:stretch");
                    w.startElement("a:fillRect");
                    w.endElement();
                    w.endElement(); // a:stretch
                    w.endElement(); // xdr:blipFill
                    
                    // 形状属性
                    w.startElement("xdr:spPr");
                    w.startElement("a:xfrm");
                    w.startElement("a:off");
                    w.writeAttribute("x", "0");
                    w.writeAttribute("y", "0");
                    w.endElement();
                    w.startElement("a:ext");
                    // 确保最小尺寸为100x100像素，转换为EMU（1像素 = 9525 EMU）
                    int64_t width_emu = static_cast<int64_t>(std::max(100.0, anchor.width) * 9525);
                    int64_t height_emu = static_cast<int64_t>(std::max(100.0, anchor.height) * 9525);
                    w.writeAttribute("cx", std::to_string(width_emu));
                    w.writeAttribute("cy", std::to_string(height_emu));
                    w.endElement();
                    w.endElement(); // a:xfrm
                    w.startElement("a:prstGeom");
                    w.writeAttribute("prst", "rect");
                    w.startElement("a:avLst");
                    w.endElement();
                    w.endElement(); // a:prstGeom
                    w.endElement(); // xdr:spPr
                    
                    w.endElement(); // xdr:pic
                    
                    // 客户端数据
                    w.startElement("xdr:clientData");
                    w.endElement();
                    
                    w.endElement(); // xdr:twoCellAnchor
                    
                    image_index++;
                }
            }
            
            w.endElement(); // xdr:wsDr
            w.flushBuffer();
            
            XML_DEBUG("Generated drawing XML with {} images", image_index);
        });
    }
};

// Drawing Relationships Generator
class DrawingRelsGenerator : public IXMLPartGenerator {
public:
    std::vector<std::string> partNames(const XMLContextView& ctx) const override {
        std::vector<std::string> parts;
        if (!ctx.workbook) return parts;
        
        // 检查所有工作表是否包含图片
        auto names = ctx.workbook->getSheetNames();
        for (size_t i = 0; i < names.size(); ++i) {
            auto ws = ctx.workbook->getSheet(static_cast<size_t>(i));
            if (ws && !ws->getImages().empty()) {
                parts.emplace_back("xl/drawings/_rels/drawing" + std::to_string(i + 1) + ".xml.rels");
            }
        }
        return parts;
    }
    
    bool generatePart(const std::string& part, const XMLContextView& ctx, IFileWriter& writer) override {
        if (!ctx.workbook) return false;
        
        // 解析绘图索引 (xl/drawings/_rels/drawing{N}.xml.rels)
        auto pos1 = part.rfind("drawing");
        auto pos2 = part.find(".xml.rels");
        if (pos1 == std::string::npos || pos2 == std::string::npos) return false;
        
        size_t number_start = pos1 + 7; // "drawing" 有7个字符
        if (number_start >= pos2) return false;
        
        std::string number_str = part.substr(number_start, pos2 - number_start);
        if (number_str.empty()) return false;
        
        int idx;
        try {
            idx = std::stoi(number_str) - 1;
        } catch (const std::exception&) {
            XML_ERROR("Failed to parse drawing index from rels path: {}, extracted: '{}'", part, number_str);
            return false;
        }
        
        auto ws = ctx.workbook->getSheet(static_cast<size_t>(idx));
        if (!ws || ws->getImages().empty()) return false;
        
        // 生成绘图关系XML
        return writeWithCallback(writer, part, [&](auto& cb){
            XMLStreamWriter w(cb);
            w.startDocument();
            w.startElement("Relationships");
            w.writeAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
            
            const auto& images = ws->getImages();
            for (size_t i = 0; i < images.size(); ++i) {
                w.startElement("Relationship");
                w.writeAttribute("Id", ("rId" + std::to_string(i + 1)).c_str());
                w.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
                w.writeAttribute("Target", ("../media/image" + std::to_string(i + 1) + ".png").c_str());
                w.endElement();
            }
            
            w.endElement();
            w.flushBuffer();
        });
    }
};

// Media Files Generator
class MediaFilesGenerator : public IXMLPartGenerator {
public:
    std::vector<std::string> partNames(const XMLContextView& ctx) const override {
        std::vector<std::string> parts;
        if (!ctx.workbook) return parts;
        
        // 收集所有工作表中的图片
        auto names = ctx.workbook->getSheetNames();
        size_t image_counter = 1;
        for (size_t i = 0; i < names.size(); ++i) {
            auto ws = ctx.workbook->getSheet(static_cast<size_t>(i));
            if (ws) {
                const auto& images = ws->getImages();
                for (const auto& image : images) {
                    std::string ext = ".png"; // 默认扩展名
                    if (image->getFormat() == core::ImageFormat::JPEG) {
                        ext = ".jpg";
                    }
                    parts.emplace_back("xl/media/image" + std::to_string(image_counter++) + ext);
                }
            }
        }
        return parts;
    }
    
    bool generatePart(const std::string& part, const XMLContextView& ctx, IFileWriter& writer) override {
        if (!ctx.workbook) return false;
        
        // 解析图片索引 (xl/media/image{N}.{ext})
        auto pos1 = part.rfind("image");
        auto pos2 = part.rfind(".");
        if (pos1 == std::string::npos || pos2 == std::string::npos || pos2 <= pos1) return false;
        
        size_t number_start = pos1 + 5; // "image" 有5个字符
        if (number_start >= pos2) return false;
        
        std::string number_str = part.substr(number_start, pos2 - number_start);
        if (number_str.empty()) return false;
        
        int target_idx;
        try {
            target_idx = std::stoi(number_str);
        } catch (const std::exception&) {
            XML_ERROR("Failed to parse image index from path: {}, extracted: '{}'", part, number_str);
            return false;
        }
        
        // 查找对应的图片
        auto names = ctx.workbook->getSheetNames();
        size_t image_counter = 1;
        for (size_t i = 0; i < names.size(); ++i) {
            auto ws = ctx.workbook->getSheet(static_cast<size_t>(i));
            if (ws) {
                const auto& images = ws->getImages();
                for (const auto& image : images) {
                    if (static_cast<int>(image_counter) == target_idx) {
                        // 找到目标图片，写入文件
                        const auto& data = image->getData();
                        // 🔧 关键修复：使用二进制数据写入，不要转换为字符串
                        if (!writer.openStreamingFile(part)) {
                            XML_ERROR("Failed to open streaming file for image: {}", part);
                            return false;
                        }
                        // 需要将uint8_t*转换为const char*
                        bool success = writer.writeStreamingChunk(reinterpret_cast<const char*>(data.data()), data.size());
                        if (!writer.closeStreamingFile()) {
                            XML_ERROR("Failed to close streaming file for image: {}", part);
                            return false;
                        }
                        return success;
                    }
                    image_counter++;
                }
            }
        }
        
        XML_ERROR("Image not found for path: {}", part);
        return false;
    }
};

// 注册到 UnifiedXMLGenerator
UnifiedXMLGenerator::UnifiedXMLGenerator(const GenerationContext& context) : context_(context) {
    registerDefaultParts();
}

UnifiedXMLGenerator::~UnifiedXMLGenerator() = default;

void UnifiedXMLGenerator::registerDefaultParts() {
    parts_.push_back(std::make_unique<ContentTypesGenerator>());
    parts_.push_back(std::make_unique<RootRelsGenerator>());
    parts_.push_back(std::make_unique<DocPropsGenerator>());
    parts_.push_back(std::make_unique<StylesGenerator>());
    parts_.push_back(std::make_unique<SharedStringsGenerator>());
    parts_.push_back(std::make_unique<ThemeGenerator>());
    parts_.push_back(std::make_unique<WorkbookPartGenerator>());
    parts_.push_back(std::make_unique<WorksheetsGenerator>());
    parts_.push_back(std::make_unique<WorksheetRelsGenerator>());
    parts_.push_back(std::make_unique<DrawingPartGenerator>());
    parts_.push_back(std::make_unique<DrawingRelsGenerator>());
    parts_.push_back(std::make_unique<MediaFilesGenerator>());
}

bool UnifiedXMLGenerator::generateAll(IFileWriter& writer) {
    XMLContextView view{};
    view.workbook = context_.workbook;
    view.format_repo = context_.format_repo;
    view.sst = context_.sst;
    view.theme = context_.workbook ? context_.workbook->getTheme() : nullptr;

    for (auto& p : parts_) {
        auto names = p->partNames(view);
        for (auto& name : names) {
            if (!p->generatePart(name, view, writer)) {
                return false;
            }
        }
    }
    return true;
}

bool UnifiedXMLGenerator::generateParts(IFileWriter& writer,
                                        const std::vector<std::string>& parts_to_generate) {
    XMLContextView view{};
    view.workbook = context_.workbook;
    view.format_repo = context_.format_repo;
    view.sst = context_.sst;
    view.theme = context_.workbook ? context_.workbook->getTheme() : nullptr;

    for (const auto& target : parts_to_generate) {
        bool handled = false;
        for (auto& p : parts_) {
            // 查询该生成器可生成的部件集合
            auto names = p->partNames(view);
            for (const auto& n : names) {
                if (n == target) {
                    if (!p->generatePart(target, view, writer)) return false;
                    handled = true;
                    break;
                }
            }
            if (handled) break;
        }
        if (!handled) {
            return false;
        }
    }
    return true;
}

}} // namespace
