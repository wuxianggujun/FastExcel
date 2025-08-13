#pragma once

#include "fastexcel/core/Image.hpp"
#include "fastexcel/xml/XMLStreamWriter.hpp"
#include "fastexcel/xml/Relationships.hpp"
#include <string>
#include <vector>
#include <memory>
#include <functional>

namespace fastexcel {
namespace xml {

/**
 * @brief Excel绘图XML生成器
 * 
 * 负责生成Excel中的绘图相关XML文件：
 * - xl/drawings/drawing1.xml - 绘图对象定义
 * - xl/drawings/_rels/drawing1.xml.rels - 绘图关系
 */
class DrawingXMLGenerator {
private:
    const std::vector<std::unique_ptr<core::Image>>* images_;
    int drawing_id_;
    
public:
    /**
     * @brief 构造函数
     * @param images 图片列表
     * @param drawing_id 绘图ID（通常与工作表ID相同）
     */
    explicit DrawingXMLGenerator(const std::vector<std::unique_ptr<core::Image>>* images, 
                                int drawing_id = 1);
    
    /**
     * @brief 析构函数
     */
    ~DrawingXMLGenerator() = default;
    
    // 禁用拷贝构造和赋值
    DrawingXMLGenerator(const DrawingXMLGenerator&) = delete;
    DrawingXMLGenerator& operator=(const DrawingXMLGenerator&) = delete;
    
    // 允许移动构造和赋值
    DrawingXMLGenerator(DrawingXMLGenerator&&) = default;
    DrawingXMLGenerator& operator=(DrawingXMLGenerator&&) = default;
    
    /**
     * @brief 生成绘图XML内容
     * @param callback 数据写入回调函数
     * @param forceGenerate 强制生成XML（即使hasImages()返回false）
     */
    void generateDrawingXML(const std::function<void(const char*, size_t)>& callback, 
                           bool forceGenerate = false) const;
    
    /**
     * @brief 生成绘图XML文件
     * @param filename 输出文件名
     */
    void generateDrawingXMLToFile(const std::string& filename) const;
    
    /**
     * @brief 生成绘图关系XML
     * @param callback 数据写入回调函数
     */
    void generateDrawingRelsXML(const std::function<void(const char*, size_t)>& callback) const;
    
    /**
     * @brief 生成绘图关系XML文件
     * @param filename 输出文件名
     */
    void generateDrawingRelsXMLToFile(const std::string& filename) const;
    
    /**
     * @brief 为工作表添加绘图关系
     * @param relationships 工作表关系对象
     * @return 绘图关系ID
     */
    std::string addWorksheetDrawingRelationship(Relationships& relationships) const;
    
    /**
     * @brief 检查是否有图片需要生成绘图XML
     * @return 是否有图片
     */
    bool hasImages() const;
    
    /**
     * @brief 获取绘图ID
     * @return 绘图ID
     */
    int getDrawingId() const { return drawing_id_; }
    
private:
    /**
     * @brief 生成单个图片的XML
     * @param writer XML写入器
     * @param image 图片对象
     * @param image_index 图片索引（从0开始）
     */
    void generateImageXML(XMLStreamWriter& writer, const core::Image& image, int image_index) const;
    
    /**
     * @brief 生成锚定信息XML
     * @param writer XML写入器
     * @param anchor 锚定信息
     */
    void generateAnchorXML(XMLStreamWriter& writer, const core::ImageAnchor& anchor) const;
    
    /**
     * @brief 生成绝对锚定XML
     * @param writer XML写入器
     * @param anchor 锚定信息
     */
    void generateAbsoluteAnchorXML(XMLStreamWriter& writer, const core::ImageAnchor& anchor) const;
    
    /**
     * @brief 生成单元格锚定XML
     * @param writer XML写入器
     * @param anchor 锚定信息
     */
    void generateOneCellAnchorXML(XMLStreamWriter& writer, const core::ImageAnchor& anchor) const;
    
    /**
     * @brief 生成双单元格锚定XML
     * @param writer XML写入器
     * @param anchor 锚定信息
     */
    void generateTwoCellAnchorXML(XMLStreamWriter& writer, const core::ImageAnchor& anchor) const;
    
    /**
     * @brief 生成图片元素XML
     * @param writer XML写入器
     * @param image 图片对象
     * @param image_index 图片索引
     */
    void generatePictureXML(XMLStreamWriter& writer, const core::Image& image, int image_index) const;
    
    /**
     * @brief 将像素转换为EMU（English Metric Units）
     * @param pixels 像素值
     * @return EMU值
     */
    static int64_t pixelsToEMU(double pixels);
    
    /**
     * @brief 将行列坐标转换为EMU偏移
     * @param row 行号
     * @param col 列号
     * @return EMU偏移值
     */
    static std::pair<int64_t, int64_t> cellToEMU(int row, int col);
    
    /**
     * @brief 生成单元格引用字符串
     * @param row 行号（0-based）
     * @param col 列号（0-based）
     * @return 单元格引用（如"A1"）
     */
    static std::string cellReference(int row, int col);
};

/**
 * @brief DrawingXMLGenerator工厂类
 */
class DrawingXMLGeneratorFactory {
public:
    /**
     * @brief 创建绘图XML生成器
     * @param images 图片列表
     * @param drawing_id 绘图ID
     * @return 生成器对象
     */
    static std::unique_ptr<DrawingXMLGenerator> create(
        const std::vector<std::unique_ptr<core::Image>>* images, 
        int drawing_id = 1);
};

}} // namespace fastexcel::xml
