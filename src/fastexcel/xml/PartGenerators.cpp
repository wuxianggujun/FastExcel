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
            w.startElement("Default"); w.writeAttribute("Extension", "gif"); w.writeAttribute("ContentType", "image/gif"); w.endElement();
            w.startElement("Default"); w.writeAttribute("Extension", "bmp"); w.writeAttribute("ContentType", "image/bmp"); w.endElement();
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
            XML_DEBUG("No worksheet or no images for drawing index {}", idx);
            return false;
        }
        
        // 🔧 关键修复：使用DrawingXMLGenerator而非硬编码XML
        const auto& images = ws->getImages();
        DrawingXMLGenerator gen(&images, idx + 1);
        XML_DEBUG("Generating drawing XML for {} images using DrawingXMLGenerator", images.size());
        
        return writeWithCallback(writer, part, [&gen](auto& cb){
            gen.generateDrawingXML(cb, true); // 强制生成，因为已经检查过图片存在
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
                
                // 🔧 关键修复：使用动态扩展名而非硬编码.png
                std::string ext = images[i]->getFileExtension();
                if (ext.empty()) {
                    XML_ERROR("Image {} has unknown format, using png as fallback", i + 1);
                    ext = "png";
                }
                std::string target = "../media/image" + std::to_string(i + 1) + "." + ext;
                w.writeAttribute("Target", target.c_str());
                w.endElement();
                
                XML_DEBUG("Added drawing relationship: rId{} -> {}", i + 1, target);
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
                // 🔧 关键修复：使用动态扩展名
                std::string ext = image->getFileExtension();
                if (ext.empty()) {
                        XML_ERROR("Image has unknown format, skipping: {}", image->getId());
                    continue;
                }
                parts.emplace_back("xl/media/image" + std::to_string(image_counter++) + "." + ext);
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
