#pragma once

#include <string>
#include <vector>
#include <memory>
#include <cstdint>

namespace fastexcel {
namespace core {

/**
 * @brief 图片格式枚举
 */
enum class ImageFormat : uint8_t {
    PNG = 0,
    JPEG = 1,
    GIF = 2,
    BMP = 3,
    UNKNOWN = 255
};

/**
 * @brief 图片锚定类型
 */
enum class ImageAnchorType : uint8_t {
    Absolute = 0,    // 绝对定位（基于像素坐标）
    OneCell = 1,     // 单元格锚定（固定大小，不随单元格缩放）
    TwoCell = 2      // 双单元格锚定（随单元格缩放）
};

/**
 * @brief 图片锚定信息
 */
struct ImageAnchor {
    ImageAnchorType type = ImageAnchorType::OneCell;
    
    // 单元格坐标（0-based）
    int from_row = 0;
    int from_col = 0;
    int to_row = 0;
    int to_col = 0;
    
    // 像素偏移和尺寸
    double offset_x = 0.0;      // 相对于锚定点的X偏移（像素）
    double offset_y = 0.0;      // 相对于锚定点的Y偏移（像素）
    double width = 0.0;         // 图片宽度（像素）
    double height = 0.0;        // 图片高度（像素）
    
    // 绝对定位坐标（仅当type为Absolute时使用）
    double abs_x = 0.0;         // 绝对X坐标（像素）
    double abs_y = 0.0;         // 绝对Y坐标（像素）
    
    ImageAnchor() = default;
    
    // 单元格锚定构造函数
    ImageAnchor(int row, int col, double w, double h, double ox = 0.0, double oy = 0.0)
        : type(ImageAnchorType::OneCell), from_row(row), from_col(col), 
          offset_x(ox), offset_y(oy), width(w), height(h) {}
    
    // 双单元格锚定构造函数
    ImageAnchor(int from_r, int from_c, int to_r, int to_c)
        : type(ImageAnchorType::TwoCell), from_row(from_r), from_col(from_c),
          to_row(to_r), to_col(to_c) {}
    
    // 绝对定位构造函数
    ImageAnchor(double x, double y, double w, double h)
        : type(ImageAnchorType::Absolute), abs_x(x), abs_y(y), width(w), height(h) {}
};

/**
 * @brief 图片类 - 管理图片数据和属性
 */
class Image {
private:
    std::string id_;                    // 图片唯一标识符
    std::string name_;                  // 图片名称
    std::string description_;           // 图片描述
    std::vector<uint8_t> data_;         // 图片二进制数据
    ImageFormat format_;                // 图片格式
    ImageAnchor anchor_;                // 锚定信息
    std::string original_filename_;     // 原始文件名
    
    // 图片尺寸信息（从图片数据中解析）
    int original_width_ = 0;
    int original_height_ = 0;
    
public:
    /**
     * @brief 默认构造函数
     */
    Image();
    
    /**
     * @brief 析构函数
     */
    ~Image() = default;
    
    // 禁用拷贝构造和赋值（使用移动语义）
    Image(const Image&) = delete;
    Image& operator=(const Image&) = delete;
    
    // 允许移动构造和赋值
    Image(Image&&) = default;
    Image& operator=(Image&&) = default;

    /**
     * @brief 深拷贝当前图片对象
     * @return 复制后的新图片对象（包含ID、锚定、元数据与二进制数据）
     */
    std::unique_ptr<Image> clone() const;
    
    /**
     * @brief 从文件创建图片
     * @param filepath 图片文件路径
     * @return 图片对象的unique_ptr，失败时返回nullptr
     */
    static std::unique_ptr<Image> fromFile(const std::string& filepath);
    
    /**
     * @brief 从内存数据创建图片
     * @param data 图片二进制数据
     * @param format 图片格式
     * @param filename 原始文件名（可选）
     * @return 图片对象的unique_ptr，失败时返回nullptr
     */
    static std::unique_ptr<Image> fromData(const std::vector<uint8_t>& data, 
                                          ImageFormat format,
                                          const std::string& filename = "");
    
    /**
     * @brief 从内存数据创建图片（移动语义）
     * @param data 图片二进制数据
     * @param format 图片格式
     * @param filename 原始文件名（可选）
     * @return 图片对象的unique_ptr，失败时返回nullptr
     */
    static std::unique_ptr<Image> fromData(std::vector<uint8_t>&& data, 
                                          ImageFormat format,
                                          const std::string& filename = "");
    
    // 属性访问器
    
    /**
     * @brief 获取图片ID
     */
    const std::string& getId() const { return id_; }
    
    /**
     * @brief 设置图片ID
     */
    void setId(const std::string& id) { id_ = id; }
    
    /**
     * @brief 获取图片名称
     */
    const std::string& getName() const { return name_; }
    
    /**
     * @brief 设置图片名称
     */
    void setName(const std::string& name) { name_ = name; }
    
    /**
     * @brief 获取图片描述
     */
    const std::string& getDescription() const { return description_; }
    
    /**
     * @brief 设置图片描述
     */
    void setDescription(const std::string& description) { description_ = description; }
    
    /**
     * @brief 获取图片格式
     */
    ImageFormat getFormat() const { return format_; }
    
    /**
     * @brief 获取图片数据
     */
    const std::vector<uint8_t>& getData() const { return data_; }
    
    /**
     * @brief 获取图片数据大小
     */
    size_t getDataSize() const { return data_.size(); }
    
    /**
     * @brief 获取原始文件名
     */
    const std::string& getOriginalFilename() const { return original_filename_; }
    
    /**
     * @brief 获取原始图片宽度
     */
    int getOriginalWidth() const { return original_width_; }
    
    /**
     * @brief 获取原始图片高度
     */
    int getOriginalHeight() const { return original_height_; }
    
    // 锚定信息管理
    
    /**
     * @brief 获取锚定信息
     */
    const ImageAnchor& getAnchor() const { return anchor_; }
    
    /**
     * @brief 设置锚定信息
     */
    void setAnchor(const ImageAnchor& anchor) { anchor_ = anchor; }
    
    /**
     * @brief 设置单元格锚定
     * @param row 行号（0-based）
     * @param col 列号（0-based）
     * @param width 图片宽度（像素）
     * @param height 图片高度（像素）
     * @param offset_x X偏移（像素）
     * @param offset_y Y偏移（像素）
     */
    void setCellAnchor(int row, int col, double width, double height, 
                      double offset_x = 0.0, double offset_y = 0.0);
    
    /**
     * @brief 设置双单元格锚定
     * @param from_row 起始行号
     * @param from_col 起始列号
     * @param to_row 结束行号
     * @param to_col 结束列号
     */
    void setRangeAnchor(int from_row, int from_col, int to_row, int to_col);
    
    /**
     * @brief 设置绝对定位
     * @param x 绝对X坐标（像素）
     * @param y 绝对Y坐标（像素）
     * @param width 图片宽度（像素）
     * @param height 图片高度（像素）
     */
    void setAbsoluteAnchor(double x, double y, double width, double height);
    
    // 工具方法
    
    /**
     * @brief 获取图片格式的文件扩展名
     */
    std::string getFileExtension() const;
    
    /**
     * @brief 获取图片格式的MIME类型
     */
    std::string getMimeType() const;
    
    /**
     * @brief 检查图片数据是否有效
     */
    bool isValid() const;
    
    /**
     * @brief 获取内存使用量
     */
    size_t getMemoryUsage() const;
    
private:
    /**
     * @brief 从图片数据中检测格式
     */
    static ImageFormat detectFormat(const std::vector<uint8_t>& data);
    
    /**
     * @brief 从图片数据中解析尺寸信息
     */
    bool parseImageDimensions();
    
    /**
     * @brief 生成唯一ID
     */
    std::string generateId() const;
};

/**
 * @brief 图片格式工具函数
 */
class ImageUtils {
public:
    /**
     * @brief 格式枚举转字符串
     */
    static std::string formatToString(ImageFormat format);
    
    /**
     * @brief 字符串转格式枚举
     */
    static ImageFormat stringToFormat(const std::string& format_str);
    
    /**
     * @brief 从文件扩展名推断格式
     */
    static ImageFormat formatFromExtension(const std::string& filename);
    
    /**
     * @brief 获取格式对应的文件扩展名
     */
    static std::string getExtension(ImageFormat format);
    
    /**
     * @brief 获取格式对应的MIME类型
     */
    static std::string getMimeType(ImageFormat format);
};

}} // namespace fastexcel::core
