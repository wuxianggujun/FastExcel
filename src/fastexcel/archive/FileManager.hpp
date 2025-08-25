#pragma once

#include "fastexcel/archive/ZipArchive.hpp"
#include "fastexcel/core/Path.hpp"
#include "fastexcel/core/Image.hpp"
#include <string>
#include <memory>
#include <unordered_map>
#include <vector>

namespace fastexcel {
namespace archive {

class FileManager {
private:
    std::unique_ptr<ZipArchive> archive_;
    std::string filename_;  // 保留用于日志
    core::Path filepath_;   // 用于实际文件操作
    
public:
    explicit FileManager(const core::Path& path);
    ~FileManager();
    
    // 文件操作
    bool open(bool create = true);
    bool close();
    
    // 写入文件
    bool writeFile(const std::string& internal_path, const std::string& content);
    bool writeFile(const std::string& internal_path, const std::vector<uint8_t>& data);
    
    // 批量写入文件 - 高性能模式
    bool writeFiles(const std::vector<std::pair<std::string, std::string>>& files);
    bool writeFiles(std::vector<std::pair<std::string, std::string>>&& files); // 移动语义版本
    
    // 流式写入文件 - 极致性能模式，直接写入ZIP
    bool openStreamingFile(const std::string& internal_path);
    bool writeStreamingChunk(const void* data, size_t size);
    bool writeStreamingChunk(const std::string& data);
    bool closeStreamingFile();
    
    // 读取文件
    bool readFile(const std::string& internal_path, std::string& content);
    bool readFile(const std::string& internal_path, std::vector<uint8_t>& data);
    
    // 检查文件是否存在
    bool fileExists(const std::string& internal_path) const;
    
    // 获取文件列表
    std::vector<std::string> listFiles() const;
    
    // 获取状态
    bool isOpen() const { return archive_ && archive_->isOpen(); }
    
    // 压缩设置
    bool setCompressionLevel(int level);
    
    // 从现有包中复制未修改的条目（编辑模式保真写回）
    // skip_prefixes: 以这些前缀开头的路径将被跳过，不复制（因为将由新的生成逻辑覆盖）
    bool copyFromExistingPackage(const core::Path& source_package,
                                 const std::vector<std::string>& skip_prefixes);
    
    // 图片文件管理
    
    /**
     * @brief 添加图片文件到媒体目录
     * @param image_id 图片ID
     * @param image_data 图片二进制数据
     * @param format 图片格式
     * @return 是否成功添加
     */
    bool addImageFile(const std::string& image_id,
                     const std::vector<uint8_t>& image_data,
                     core::ImageFormat format);
    
    /**
     * @brief 添加图片文件到媒体目录（移动语义）
     * @param image_id 图片ID
     * @param image_data 图片二进制数据
     * @param format 图片格式
     * @return 是否成功添加
     */
    bool addImageFile(const std::string& image_id,
                     std::vector<uint8_t>&& image_data,
                     core::ImageFormat format);
    
    /**
     * @brief 从Image对象添加图片文件
     * @param image 图片对象
     * @return 是否成功添加
     */
    bool addImageFile(const core::Image& image);
    
    /**
     * @brief 批量添加图片文件
     * @param images 图片列表
     * @return 成功添加的图片数量
     */
    int addImageFiles(const std::vector<std::unique_ptr<core::Image>>& images);
    
    /**
     * @brief 添加绘图XML文件
     * @param drawing_id 绘图ID
     * @param xml_content XML内容
     * @return 是否成功添加
     */
    bool addDrawingXML(int drawing_id, const std::string& xml_content);
    
    /**
     * @brief 添加绘图关系XML文件
     * @param drawing_id 绘图ID
     * @param xml_content XML内容
     * @return 是否成功添加
     */
    bool addDrawingRelsXML(int drawing_id, const std::string& xml_content);
    
    /**
     * @brief 检查媒体目录中是否存在指定图片
     * @param image_id 图片ID
     * @param format 图片格式
     * @return 是否存在
     */
    bool imageExists(const std::string& image_id, core::ImageFormat format) const;
    
    /**
     * @brief 获取图片文件的内部路径
     * @param image_id 图片ID
     * @param format 图片格式
     * @return 内部路径
     */
    static std::string getImagePath(const std::string& image_id, core::ImageFormat format);
    
    /**
     * @brief 获取绘图XML文件的内部路径
     * @param drawing_id 绘图ID
     * @return 内部路径
     */
    static std::string getDrawingPath(int drawing_id);
    
    /**
     * @brief 获取绘图关系XML文件的内部路径
     * @param drawing_id 绘图ID
     * @return 内部路径
     */
    static std::string getDrawingRelsPath(int drawing_id);
};

}} // namespace fastexcel::archive
