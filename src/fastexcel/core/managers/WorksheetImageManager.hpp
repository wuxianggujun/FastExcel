#pragma once

#include "fastexcel/core/Image.hpp"
#include <vector>
#include <memory>
#include <string>

namespace fastexcel {
namespace core {

class WorksheetImageManager {
public:
    WorksheetImageManager();
    
    // 图片插入 - 单元格锚定
    std::string insertImage(int row, int col, const std::string& image_path);
    std::string insertImage(int row, int col, std::unique_ptr<Image> image);
    
    // 图片插入 - 范围锚定
    std::string insertImage(int from_row, int from_col, int to_row, int to_col,
                           const std::string& image_path);
    std::string insertImage(int from_row, int from_col, int to_row, int to_col,
                           std::unique_ptr<Image> image);
    
    // 图片插入 - 绝对定位
    std::string insertImageAt(double x, double y, double width, double height,
                             const std::string& image_path);
    std::string insertImageAt(double x, double y, double width, double height,
                             std::unique_ptr<Image> image);
    
    // 字符串地址支持
    std::string insertImage(const std::string& address, const std::string& image_path);
    std::string insertImage(const std::string& address, std::unique_ptr<Image> image);
    std::string insertImageRange(const std::string& range, const std::string& image_path);
    std::string insertImageRange(const std::string& range, std::unique_ptr<Image> image);
    
    // 图片查找和管理
    const Image* findImage(const std::string& image_id) const;
    Image* findImage(const std::string& image_id);
    bool removeImage(const std::string& image_id);
    void clearImages();
    
    // 状态查询
    size_t getImageCount() const { return images_.size(); }
    bool hasImages() const { return !images_.empty(); }
    const std::vector<std::unique_ptr<Image>>& getImages() const { return images_; }
    
    // 内存管理
    size_t getImagesMemoryUsage() const;
    
    // 清理
    void clear();

private:
    std::vector<std::unique_ptr<Image>> images_;
    int next_image_id_ = 1;
    
    void validateCellPosition(int row, int col) const;
    void validateRange(int first_row, int first_col, int last_row, int last_col) const;
    std::string generateNextImageId();
};

} // namespace core
} // namespace fastexcel