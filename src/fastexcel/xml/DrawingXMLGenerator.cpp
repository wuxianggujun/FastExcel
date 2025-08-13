#include "fastexcel/xml/DrawingXMLGenerator.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/utils/CommonUtils.hpp"
#include <fstream>
#include <sstream>
#include <iomanip>

namespace fastexcel {
namespace xml {

// Excel中的常量定义
namespace {
    // EMU (English Metric Units) 转换常量
    constexpr int64_t PIXELS_PER_INCH = 96;
    constexpr int64_t EMU_PER_INCH = 914400;
    constexpr int64_t EMU_PER_PIXEL = EMU_PER_INCH / PIXELS_PER_INCH;
    
    // Excel默认行高和列宽（像素）
    constexpr double DEFAULT_ROW_HEIGHT_PIXELS = 20.0;
    constexpr double DEFAULT_COL_WIDTH_PIXELS = 64.0;
}

DrawingXMLGenerator::DrawingXMLGenerator(const std::vector<std::unique_ptr<core::Image>>* images, 
                                        int drawing_id)
    : images_(images), drawing_id_(drawing_id) {
}

void DrawingXMLGenerator::generateDrawingXML(const std::function<void(const char*, size_t)>& callback, 
                                             bool forceGenerate) const {
    if (!forceGenerate && !hasImages()) {
        FASTEXCEL_LOG_DEBUG("No images to generate drawing XML");
        return;
    }
    
    XMLStreamWriter writer(callback);
    
    // XML声明
    writer.startDocument();
    
    // 根元素
    writer.startElement("xdr:wsDr");
    writer.writeAttribute("xmlns:xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
    writer.writeAttribute("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main");
    writer.writeAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    
    // 生成每个图片的XML
    int image_index = 0;
    if (images_) {
        for (const auto& image : *images_) {
            if (image && image->isValid()) {
                generateImageXML(writer, *image, image_index++);
            }
        }
    }
    
    writer.endElement(); // xdr:wsDr
    
    // 🔧 关键修复：保证XML数据被完整写入
    writer.flushBuffer();
    
    FASTEXCEL_LOG_DEBUG("Generated drawing XML with {} images", image_index);
}

void DrawingXMLGenerator::generateDrawingXMLToFile(const std::string& filename) const {
    std::ofstream file(filename, std::ios::binary);
    if (!file.is_open()) {
        FASTEXCEL_LOG_ERROR("Failed to open file for writing: {}", filename);
        return;
    }
    
    generateDrawingXML([&file](const char* data, size_t size) {
        file.write(data, size);
    });
    
    file.close();
    FASTEXCEL_LOG_DEBUG("Drawing XML written to file: {}", filename);
}

void DrawingXMLGenerator::generateDrawingRelsXML(const std::function<void(const char*, size_t)>& callback) const {
    if (!hasImages()) {
        FASTEXCEL_LOG_DEBUG("No images to generate drawing relationships XML");
        return;
    }
    
    Relationships relationships;
    
    // 为每个图片添加关系
    int image_index = 1;
    for (const auto& image : *images_) {
        if (image && image->isValid()) {
            std::string rel_id = "rId" + std::to_string(image_index);
            // 修复：使用统一的图片文件命名规则 image{index}.{ext}
            std::string ext = image->getFileExtension();
            std::string target = "../media/image" + std::to_string(image_index) + "." + ext;
            
            relationships.addRelationship(
                rel_id,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                target
            );
            
            image_index++;
        }
    }
    
    // 生成关系XML
    relationships.generate(callback);
    
    FASTEXCEL_LOG_DEBUG("Generated drawing relationships XML with {} images", image_index - 1);
}

void DrawingXMLGenerator::generateDrawingRelsXMLToFile(const std::string& filename) const {
    std::ofstream file(filename, std::ios::binary);
    if (!file.is_open()) {
        FASTEXCEL_LOG_ERROR("Failed to open file for writing: {}", filename);
        return;
    }
    
    generateDrawingRelsXML([&file](const char* data, size_t size) {
        file.write(data, size);
    });
    
    file.close();
    FASTEXCEL_LOG_DEBUG("Drawing relationships XML written to file: {}", filename);
}

std::string DrawingXMLGenerator::addWorksheetDrawingRelationship(Relationships& relationships) const {
    if (!hasImages()) {
        return "";
    }
    
    std::string drawing_target = "../drawings/drawing" + std::to_string(drawing_id_) + ".xml";
    return relationships.addAutoRelationship(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
        drawing_target
    );
}

bool DrawingXMLGenerator::hasImages() const {
    if (!images_ || images_->empty()) {
        return false;
    }
    
    // 检查是否有有效的图片
    for (const auto& image : *images_) {
        if (image && image->isValid()) {
            return true;
        }
    }
    
    return false;
}

void DrawingXMLGenerator::generateImageXML(XMLStreamWriter& writer, const core::Image& image, int image_index) const {
    const auto& anchor = image.getAnchor();
    
    // 根据锚定类型生成不同的XML
    switch (anchor.type) {
        case core::ImageAnchorType::Absolute:
            generateAbsoluteAnchorXML(writer, anchor);
            break;
        case core::ImageAnchorType::OneCell:
            generateOneCellAnchorXML(writer, anchor);
            break;
        case core::ImageAnchorType::TwoCell:
            generateTwoCellAnchorXML(writer, anchor);
            break;
    }
    
    // 在锚定元素内生成图片元素
    generatePictureXML(writer, image, image_index);
    
    // 客户端数据（可选）
    writer.startElement("xdr:clientData");
    writer.endElement(); // xdr:clientData
    
    // 结束锚定元素
    switch (anchor.type) {
        case core::ImageAnchorType::Absolute:
            writer.endElement(); // xdr:absoluteAnchor
            break;
        case core::ImageAnchorType::OneCell:
            writer.endElement(); // xdr:oneCellAnchor
            break;
        case core::ImageAnchorType::TwoCell:
            writer.endElement(); // xdr:twoCellAnchor
            break;
    }
}

void DrawingXMLGenerator::generateAbsoluteAnchorXML(XMLStreamWriter& writer, const core::ImageAnchor& anchor) const {
    writer.startElement("xdr:absoluteAnchor");
    
    // 位置
    writer.startElement("xdr:pos");
    writer.writeAttribute("x", std::to_string(pixelsToEMU(anchor.abs_x)));
    writer.writeAttribute("y", std::to_string(pixelsToEMU(anchor.abs_y)));
    writer.endElement(); // xdr:pos
    
    // 尺寸
    writer.startElement("xdr:ext");
    writer.writeAttribute("cx", std::to_string(pixelsToEMU(anchor.width)));
    writer.writeAttribute("cy", std::to_string(pixelsToEMU(anchor.height)));
    writer.endElement(); // xdr:ext
}

void DrawingXMLGenerator::generateOneCellAnchorXML(XMLStreamWriter& writer, const core::ImageAnchor& anchor) const {
    writer.startElement("xdr:oneCellAnchor");
    
    // 起始位置
    writer.startElement("xdr:from");
    writer.startElement("xdr:col");
    writer.writeText(std::to_string(anchor.from_col));
    writer.endElement();
    writer.startElement("xdr:colOff");
    writer.writeText(std::to_string(pixelsToEMU(anchor.offset_x)));
    writer.endElement();
    writer.startElement("xdr:row");
    writer.writeText(std::to_string(anchor.from_row));
    writer.endElement();
    writer.startElement("xdr:rowOff");
    writer.writeText(std::to_string(pixelsToEMU(anchor.offset_y)));
    writer.endElement();
    writer.endElement(); // xdr:from
    
    // 尺寸
    writer.startElement("xdr:ext");
    writer.writeAttribute("cx", std::to_string(pixelsToEMU(anchor.width)));
    writer.writeAttribute("cy", std::to_string(pixelsToEMU(anchor.height)));
    writer.endElement(); // xdr:ext
}

void DrawingXMLGenerator::generateTwoCellAnchorXML(XMLStreamWriter& writer, const core::ImageAnchor& anchor) const {
    writer.startElement("xdr:twoCellAnchor");
    
    // 起始位置
    writer.startElement("xdr:from");
    writer.startElement("xdr:col");
    writer.writeText(std::to_string(anchor.from_col));
    writer.endElement();
    writer.startElement("xdr:colOff");
    writer.writeText(std::to_string(pixelsToEMU(anchor.offset_x)));
    writer.endElement();
    writer.startElement("xdr:row");
    writer.writeText(std::to_string(anchor.from_row));
    writer.endElement();
    writer.startElement("xdr:rowOff");
    writer.writeText(std::to_string(pixelsToEMU(anchor.offset_y)));
    writer.endElement();
    writer.endElement(); // xdr:from
    
    // 结束位置
    writer.startElement("xdr:to");
    writer.startElement("xdr:col");
    writer.writeText(std::to_string(anchor.to_col));
    writer.endElement();
    writer.startElement("xdr:colOff");
    writer.writeText("0");
    writer.endElement();
    writer.startElement("xdr:row");
    writer.writeText(std::to_string(anchor.to_row));
    writer.endElement();
    writer.startElement("xdr:rowOff");
    writer.writeText("0");
    writer.endElement();
    writer.endElement(); // xdr:to
}

void DrawingXMLGenerator::generatePictureXML(XMLStreamWriter& writer, const core::Image& image, int image_index) const {
    writer.startElement("xdr:pic");
    
    // 非可视属性
    writer.startElement("xdr:nvPicPr");
    
    // 非可视绘图属性
    writer.startElement("xdr:cNvPr");
    writer.writeAttribute("id", std::to_string(image_index + 2)); // ID从2开始
    writer.writeAttribute("name", image.getName().empty() ? image.getId() : image.getName());
    if (!image.getDescription().empty()) {
        writer.writeAttribute("descr", image.getDescription());
    }
    writer.endElement(); // xdr:cNvPr
    
    // 非可视图片属性
    writer.startElement("xdr:cNvPicPr");
    writer.startElement("a:picLocks");
    writer.writeAttribute("noChangeAspect", "1");
    writer.endElement(); // a:picLocks
    writer.endElement(); // xdr:cNvPicPr
    
    writer.endElement(); // xdr:nvPicPr
    
    // 图片填充
    writer.startElement("xdr:blipFill");
    
    writer.startElement("a:blip");
    writer.writeAttribute("r:embed", "rId" + std::to_string(image_index + 1));
    writer.endElement(); // a:blip
    
    writer.startElement("a:stretch");
    writer.startElement("a:fillRect");
    writer.endElement();
    writer.endElement(); // a:stretch
    
    writer.endElement(); // xdr:blipFill
    
    // 形状属性
    writer.startElement("xdr:spPr");
    
    writer.startElement("a:xfrm");
    writer.startElement("a:off");
    writer.writeAttribute("x", "0");
    writer.writeAttribute("y", "0");
    writer.endElement(); // a:off
    writer.startElement("a:ext");
    writer.writeAttribute("cx", std::to_string(pixelsToEMU(image.getAnchor().width)));
    writer.writeAttribute("cy", std::to_string(pixelsToEMU(image.getAnchor().height)));
    writer.endElement(); // a:ext
    writer.endElement(); // a:xfrm
    
    writer.startElement("a:prstGeom");
    writer.writeAttribute("prst", "rect");
    writer.startElement("a:avLst");
    writer.endElement();
    writer.endElement(); // a:prstGeom
    
    writer.endElement(); // xdr:spPr
    
    writer.endElement(); // xdr:pic
}

int64_t DrawingXMLGenerator::pixelsToEMU(double pixels) {
    return static_cast<int64_t>(pixels * EMU_PER_PIXEL);
}

std::pair<int64_t, int64_t> DrawingXMLGenerator::cellToEMU(int row, int col) {
    int64_t x = static_cast<int64_t>(col * DEFAULT_COL_WIDTH_PIXELS * EMU_PER_PIXEL);
    int64_t y = static_cast<int64_t>(row * DEFAULT_ROW_HEIGHT_PIXELS * EMU_PER_PIXEL);
    return {x, y};
}

std::string DrawingXMLGenerator::cellReference(int row, int col) {
    return utils::CommonUtils::cellReference(row, col);
}

// ========== DrawingXMLGeneratorFactory实现 ==========

std::unique_ptr<DrawingXMLGenerator> DrawingXMLGeneratorFactory::create(
    const std::vector<std::unique_ptr<core::Image>>* images, 
    int drawing_id) {
    return std::make_unique<DrawingXMLGenerator>(images, drawing_id);
}

}} // namespace fastexcel::xml