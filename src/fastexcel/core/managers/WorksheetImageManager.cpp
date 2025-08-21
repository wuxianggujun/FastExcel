#include "WorksheetImageManager.hpp"
#include "fastexcel/core/Exception.hpp"
#include "fastexcel/utils/AddressParser.hpp"
#include "fastexcel/utils/CommonUtils.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <algorithm>

namespace fastexcel {
namespace core {

WorksheetImageManager::WorksheetImageManager() {
}

std::string WorksheetImageManager::insertImage(int row, int col, const std::string& image_path) {
    FASTEXCEL_LOG_DEBUG("Inserting image from file: {} at cell ({}, {})", image_path, row, col);
    
    auto image = Image::fromFile(image_path);
    if (!image) {
        FASTEXCEL_LOG_ERROR("Failed to load image from file: {}", image_path);
        return "";
    }
    
    return insertImage(row, col, std::move(image));
}

std::string WorksheetImageManager::insertImage(int row, int col, std::unique_ptr<Image> image) {
    if (!image) {
        FASTEXCEL_LOG_ERROR("Cannot insert null image");
        return "";
    }
    
    validateCellPosition(row, col);
    
    // 设置单元格锚定
    image->setCellAnchor(row, col, image->getAnchor().width, image->getAnchor().height);
    
    // 生成唯一ID
    std::string image_id = generateNextImageId();
    image->setId(image_id);
    
    // 添加到图片列表
    images_.push_back(std::move(image));
    
    FASTEXCEL_LOG_INFO("Successfully inserted image: {} at cell ({}, {})", image_id, row, col);
    return image_id;
}

std::string WorksheetImageManager::insertImage(int from_row, int from_col, int to_row, int to_col,
                                              const std::string& image_path) {
    FASTEXCEL_LOG_DEBUG("Inserting image from file: {} in range ({},{}) to ({},{})",
                       image_path, from_row, from_col, to_row, to_col);
    
    auto image = Image::fromFile(image_path);
    if (!image) {
        FASTEXCEL_LOG_ERROR("Failed to load image from file: {}", image_path);
        return "";
    }
    
    return insertImage(from_row, from_col, to_row, to_col, std::move(image));
}

std::string WorksheetImageManager::insertImage(int from_row, int from_col, int to_row, int to_col,
                                              std::unique_ptr<Image> image) {
    if (!image) {
        FASTEXCEL_LOG_ERROR("Cannot insert null image");
        return "";
    }
    
    validateRange(from_row, from_col, to_row, to_col);
    
    // 设置双单元格锚定
    image->setRangeAnchor(from_row, from_col, to_row, to_col);
    
    // 生成唯一ID
    std::string image_id = generateNextImageId();
    image->setId(image_id);
    
    // 添加到图片列表
    images_.push_back(std::move(image));
    
    FASTEXCEL_LOG_INFO("Successfully inserted image: {} in range ({},{}) to ({},{})",
                      image_id, from_row, from_col, to_row, to_col);
    return image_id;
}

std::string WorksheetImageManager::insertImageAt(double x, double y, double width, double height,
                                                const std::string& image_path) {
    FASTEXCEL_LOG_DEBUG("Inserting image from file: {} at absolute position ({}, {}) with size {}x{}",
                       image_path, x, y, width, height);
    
    auto image = Image::fromFile(image_path);
    if (!image) {
        FASTEXCEL_LOG_ERROR("Failed to load image from file: {}", image_path);
        return "";
    }
    
    return insertImageAt(x, y, width, height, std::move(image));
}

std::string WorksheetImageManager::insertImageAt(double x, double y, double width, double height,
                                                std::unique_ptr<Image> image) {
    if (!image) {
        FASTEXCEL_LOG_ERROR("Cannot insert null image");
        return "";
    }
    
    // 设置绝对定位
    image->setAbsoluteAnchor(x, y, width, height);
    
    // 生成唯一ID
    std::string image_id = generateNextImageId();
    image->setId(image_id);
    
    // 添加到图片列表
    images_.push_back(std::move(image));
    
    FASTEXCEL_LOG_INFO("Successfully inserted image: {} at absolute position ({}, {}) with size {}x{}",
                      image_id, x, y, width, height);
    return image_id;
}

std::string WorksheetImageManager::insertImage(const std::string& address, const std::string& image_path) {
    try {
        auto [sheet, row, col] = utils::AddressParser::parseAddress(address);
        return insertImage(row, col, image_path);
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Failed to parse address '{}': {}", address, e.what());
        return "";
    }
}

std::string WorksheetImageManager::insertImage(const std::string& address, std::unique_ptr<Image> image) {
    try {
        auto [sheet, row, col] = utils::AddressParser::parseAddress(address);
        return insertImage(row, col, std::move(image));
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Failed to parse address '{}': {}", address, e.what());
        return "";
    }
}

std::string WorksheetImageManager::insertImageRange(const std::string& range, const std::string& image_path) {
    try {
        auto [sheet, start_row, start_col, end_row, end_col] = utils::AddressParser::parseRange(range);
        return insertImage(start_row, start_col, end_row, end_col, image_path);
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Failed to parse range '{}': {}", range, e.what());
        return "";
    }
}

std::string WorksheetImageManager::insertImageRange(const std::string& range, std::unique_ptr<Image> image) {
    try {
        auto [sheet, start_row, start_col, end_row, end_col] = utils::AddressParser::parseRange(range);
        return insertImage(start_row, start_col, end_row, end_col, std::move(image));
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Failed to parse range '{}': {}", range, e.what());
        return "";
    }
}

const Image* WorksheetImageManager::findImage(const std::string& image_id) const {
    auto it = std::find_if(images_.begin(), images_.end(),
                          [&image_id](const std::unique_ptr<Image>& img) {
                              return img && img->getId() == image_id;
                          });
    return (it != images_.end()) ? it->get() : nullptr;
}

Image* WorksheetImageManager::findImage(const std::string& image_id) {
    auto it = std::find_if(images_.begin(), images_.end(),
                          [&image_id](const std::unique_ptr<Image>& img) {
                              return img && img->getId() == image_id;
                          });
    return (it != images_.end()) ? it->get() : nullptr;
}

bool WorksheetImageManager::removeImage(const std::string& image_id) {
    auto it = std::find_if(images_.begin(), images_.end(),
                          [&image_id](const std::unique_ptr<Image>& img) {
                              return img && img->getId() == image_id;
                          });
    
    if (it != images_.end()) {
        FASTEXCEL_LOG_INFO("Removed image: {}", image_id);
        images_.erase(it);
        return true;
    }
    
    FASTEXCEL_LOG_WARN("Image not found for removal: {}", image_id);
    return false;
}

void WorksheetImageManager::clearImages() {
    if (!images_.empty()) {
        size_t count = images_.size();
        images_.clear();
        FASTEXCEL_LOG_INFO("Cleared {} images", count);
    }
}

size_t WorksheetImageManager::getImagesMemoryUsage() const {
    size_t total_memory = 0;
    for (const auto& image : images_) {
        if (image) {
            total_memory += image->getMemoryUsage();
        }
    }
    return total_memory;
}

void WorksheetImageManager::clear() {
    clearImages();
    next_image_id_ = 1;
}

void WorksheetImageManager::validateCellPosition(int row, int col) const {
    FASTEXCEL_VALIDATE_CELL_POSITION(row, col);
}

void WorksheetImageManager::validateRange(int first_row, int first_col, int last_row, int last_col) const {
    FASTEXCEL_VALIDATE_RANGE(first_row, first_col, last_row, last_col);
}

std::string WorksheetImageManager::generateNextImageId() {
    return "img" + std::to_string(next_image_id_++);
}

} // namespace core
} // namespace fastexcel
