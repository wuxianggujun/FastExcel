#include "fastexcel/core/Image.hpp"
#include "fastexcel/utils/Logger.hpp"

// 包含stb_image头文件
#define STB_IMAGE_IMPLEMENTATION
#include "stb_image.h"
#include <fstream>
#include <sstream>
#include <iomanip>
#include <algorithm>
#include <cstring>

namespace fastexcel {
namespace core {

// Image 类实现

Image::Image() : format_(ImageFormat::UNKNOWN) {
    id_ = generateId();
}

std::unique_ptr<Image> Image::clone() const {
    // 进行字段级深拷贝，保持ID与锚定信息一致
    auto copy = std::make_unique<Image>();
    copy->id_ = id_;
    copy->name_ = name_;
    copy->description_ = description_;
    copy->data_ = data_;
    copy->format_ = format_;
    copy->anchor_ = anchor_;
    copy->original_filename_ = original_filename_;
    copy->original_width_ = original_width_;
    copy->original_height_ = original_height_;
    return copy;
}

std::unique_ptr<Image> Image::fromFile(const std::string& filepath) {
    FASTEXCEL_LOG_DEBUG("Loading image from file: {}", filepath);
    
    // 读取文件数据
    std::ifstream file(filepath, std::ios::binary);
    if (!file.is_open()) {
        FASTEXCEL_LOG_ERROR("Failed to open image file: {}", filepath);
        return nullptr;
    }
    
    // 获取文件大小
    file.seekg(0, std::ios::end);
    size_t file_size = file.tellg();
    file.seekg(0, std::ios::beg);
    
    if (file_size == 0) {
        FASTEXCEL_LOG_ERROR("Image file is empty: {}", filepath);
        return nullptr;
    }
    
    // 读取文件内容
    std::vector<uint8_t> data(file_size);
    file.read(reinterpret_cast<char*>(data.data()), file_size);
    file.close();
    
    if (file.gcount() != static_cast<std::streamsize>(file_size)) {
        FASTEXCEL_LOG_ERROR("Failed to read complete image file: {}", filepath);
        return nullptr;
    }
    
    // 从文件扩展名推断格式
    ImageFormat format = ImageUtils::formatFromExtension(filepath);
    if (format == ImageFormat::UNKNOWN) {
        // 尝试从数据检测格式
        format = detectFormat(data);
    }
    
    // 创建图片对象
    auto image = fromData(std::move(data), format, filepath);
    if (image) {
        // 提取文件名作为图片名称
        size_t pos = filepath.find_last_of("/\\");
        std::string filename = (pos != std::string::npos) ? filepath.substr(pos + 1) : filepath;
        image->setName(filename);
        
        FASTEXCEL_LOG_INFO("Successfully loaded image: {} ({}x{}, {} bytes)", 
                          filename, image->getOriginalWidth(), image->getOriginalHeight(), 
                          image->getDataSize());
    }
    
    return image;
}

std::unique_ptr<Image> Image::fromData(const std::vector<uint8_t>& data, 
                                      ImageFormat format,
                                      const std::string& filename) {
    std::vector<uint8_t> data_copy = data;
    return fromData(std::move(data_copy), format, filename);
}

std::unique_ptr<Image> Image::fromData(std::vector<uint8_t>&& data, 
                                      ImageFormat format,
                                      const std::string& filename) {
    if (data.empty()) {
        FASTEXCEL_LOG_ERROR("Image data is empty");
        return nullptr;
    }
    
    auto image = std::make_unique<Image>();
    image->data_ = std::move(data);
    image->format_ = format;
    image->original_filename_ = filename;
    
    // 如果格式未知，尝试检测
    if (image->format_ == ImageFormat::UNKNOWN) {
        image->format_ = detectFormat(image->data_);
        if (image->format_ == ImageFormat::UNKNOWN) {
            FASTEXCEL_LOG_ERROR("Unable to detect image format");
            return nullptr;
        }
    }
    
    // 解析图片尺寸
    if (!image->parseImageDimensions()) {
        FASTEXCEL_LOG_ERROR("Failed to parse image dimensions");
        return nullptr;
    }
    
    // 设置默认锚定（如果原始尺寸可用）
    if (image->original_width_ > 0 && image->original_height_ > 0) {
        image->anchor_.width = static_cast<double>(image->original_width_);
        image->anchor_.height = static_cast<double>(image->original_height_);
    }
    
    FASTEXCEL_LOG_DEBUG("Created image object: format={}, size={}x{}, data_size={}", 
                       static_cast<int>(image->format_), 
                       image->original_width_, image->original_height_, 
                       image->data_.size());
    
    return image;
}

void Image::setCellAnchor(int row, int col, double width, double height, 
                         double offset_x, double offset_y) {
    anchor_.type = ImageAnchorType::OneCell;
    anchor_.from_row = row;
    anchor_.from_col = col;
    anchor_.width = width;
    anchor_.height = height;
    anchor_.offset_x = offset_x;
    anchor_.offset_y = offset_y;
}

void Image::setRangeAnchor(int from_row, int from_col, int to_row, int to_col) {
    anchor_.type = ImageAnchorType::TwoCell;
    anchor_.from_row = from_row;
    anchor_.from_col = from_col;
    anchor_.to_row = to_row;
    anchor_.to_col = to_col;
}

void Image::setAbsoluteAnchor(double x, double y, double width, double height) {
    anchor_.type = ImageAnchorType::Absolute;
    anchor_.abs_x = x;
    anchor_.abs_y = y;
    anchor_.width = width;
    anchor_.height = height;
}

std::string Image::getFileExtension() const {
    return ImageUtils::getExtension(format_);
}

std::string Image::getMimeType() const {
    return ImageUtils::getMimeType(format_);
}

bool Image::isValid() const {
    return !data_.empty() && format_ != ImageFormat::UNKNOWN && 
           original_width_ > 0 && original_height_ > 0;
}

size_t Image::getMemoryUsage() const {
    return sizeof(Image) + data_.size() + 
           id_.capacity() + name_.capacity() + 
           description_.capacity() + original_filename_.capacity();
}

ImageFormat Image::detectFormat(const std::vector<uint8_t>& data) {
    if (data.size() < 8) {
        return ImageFormat::UNKNOWN;
    }
    
    const uint8_t* bytes = data.data();
    
    // PNG: 89 50 4E 47 0D 0A 1A 0A
    if (bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47 &&
        bytes[4] == 0x0D && bytes[5] == 0x0A && bytes[6] == 0x1A && bytes[7] == 0x0A) {
        return ImageFormat::PNG;
    }
    
    // JPEG: FF D8 FF
    if (bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF) {
        return ImageFormat::JPEG;
    }
    
    // GIF: GIF87a or GIF89a
    if (data.size() >= 6 && 
        bytes[0] == 'G' && bytes[1] == 'I' && bytes[2] == 'F' &&
        bytes[3] == '8' && (bytes[4] == '7' || bytes[4] == '9') && bytes[5] == 'a') {
        return ImageFormat::GIF;
    }
    
    // BMP: BM
    if (bytes[0] == 'B' && bytes[1] == 'M') {
        return ImageFormat::BMP;
    }
    
    return ImageFormat::UNKNOWN;
}

bool Image::parseImageDimensions() {
    if (data_.empty()) {
        return false;
    }
    
    int width, height, channels;
    
    // 使用stb_image解析图片信息（不加载像素数据）
    stbi_info_from_memory(data_.data(), static_cast<int>(data_.size()), 
                         &width, &height, &channels);
    
    if (width <= 0 || height <= 0) {
        FASTEXCEL_LOG_ERROR("Invalid image dimensions: {}x{}", width, height);
        return false;
    }
    
    original_width_ = width;
    original_height_ = height;
    
    FASTEXCEL_LOG_DEBUG("Parsed image dimensions: {}x{}, channels={}", 
                       width, height, channels);
    
    return true;
}

std::string Image::generateId() const {
    static int counter = 1;
    std::ostringstream oss;
    oss << "img" << std::setfill('0') << std::setw(6) << counter++;
    return oss.str();
}

// ImageUtils 类实现

std::string ImageUtils::formatToString(ImageFormat format) {
    switch (format) {
        case ImageFormat::PNG:  return "PNG";
        case ImageFormat::JPEG: return "JPEG";
        case ImageFormat::GIF:  return "GIF";
        case ImageFormat::BMP:  return "BMP";
        default:                return "UNKNOWN";
    }
}

ImageFormat ImageUtils::stringToFormat(const std::string& format_str) {
    std::string upper_str = format_str;
    std::transform(upper_str.begin(), upper_str.end(), upper_str.begin(), ::toupper);
    
    if (upper_str == "PNG") return ImageFormat::PNG;
    if (upper_str == "JPEG" || upper_str == "JPG") return ImageFormat::JPEG;
    if (upper_str == "GIF") return ImageFormat::GIF;
    if (upper_str == "BMP") return ImageFormat::BMP;
    
    return ImageFormat::UNKNOWN;
}

ImageFormat ImageUtils::formatFromExtension(const std::string& filename) {
    size_t dot_pos = filename.find_last_of('.');
    if (dot_pos == std::string::npos) {
        return ImageFormat::UNKNOWN;
    }
    
    std::string ext = filename.substr(dot_pos + 1);
    std::transform(ext.begin(), ext.end(), ext.begin(), ::tolower);
    
    if (ext == "png") return ImageFormat::PNG;
    if (ext == "jpg" || ext == "jpeg") return ImageFormat::JPEG;
    if (ext == "gif") return ImageFormat::GIF;
    if (ext == "bmp") return ImageFormat::BMP;
    
    return ImageFormat::UNKNOWN;
}

std::string ImageUtils::getExtension(ImageFormat format) {
    switch (format) {
        case ImageFormat::PNG:  return "png";
        case ImageFormat::JPEG: return "jpg";
        case ImageFormat::GIF:  return "gif";
        case ImageFormat::BMP:  return "bmp";
        default:                return "";
    }
}

std::string ImageUtils::getMimeType(ImageFormat format) {
    switch (format) {
        case ImageFormat::PNG:  return "image/png";
        case ImageFormat::JPEG: return "image/jpeg";
        case ImageFormat::GIF:  return "image/gif";
        case ImageFormat::BMP:  return "image/bmp";
        default:                return "application/octet-stream";
    }
}

}} // namespace fastexcel::core
