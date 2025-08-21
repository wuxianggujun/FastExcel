#pragma once

#include <string>
#include <cstdint>
#include <functional>

namespace fastexcel {
namespace core {

/**
 * @brief Color类 - Excel颜色管理
 * 
 * 提供完整的Excel颜色支持，包括：
 * - RGB颜色
 * - 主题颜色
 * - 索引颜色
 * - 自动颜色
 * - 颜色转换和格式化
 */
class Color {
public:
    // 颜色类型
    enum class Type : uint8_t {
        RGB = 0,        // RGB颜色
        Theme = 1,      // 主题颜色
        Indexed = 2,    // 索引颜色
        Auto = 3        // 自动颜色
    };

private:
    Type type_;
    uint32_t value_;    // RGB值、主题索引或颜色索引
    double tint_;       // 色调调整 (-1.0 到 1.0)

public:
    // 构造函数
    
    /**
     * @brief 默认构造函数（黑色）
     */
    Color() : type_(Type::RGB), value_(0x000000), tint_(0.0) {}
    
    /**
     * @brief RGB三参数构造函数 - 更方便的颜色创建方式
     * @param red 红色分量 (0-255)
     * @param green 绿色分量 (0-255)  
     * @param blue 蓝色分量 (0-255)
     */
    Color(uint8_t red, uint8_t green, uint8_t blue) 
        : type_(Type::RGB), value_((static_cast<uint32_t>(red) << 16) | 
                                  (static_cast<uint32_t>(green) << 8) | 
                                  static_cast<uint32_t>(blue)), tint_(0.0) {}
    
    /**
     * @brief RGB颜色构造函数
     * @param rgb RGB值
     */
    explicit Color(uint32_t rgb) : type_(Type::RGB), value_(rgb & 0xFFFFFF), tint_(0.0) {}
    
    /**
     * @brief 主题颜色构造函数 - 使用静态方法避免歧义
     */
    static Color fromTheme(uint8_t theme_index, double tint = 0.0) {
        Color color;
        color.type_ = Type::Theme;
        color.value_ = theme_index;
        color.tint_ = tint;
        return color;
    }
    
    /**
     * @brief 索引颜色构造函数
     * @param color_index 颜色索引
     */
    static Color fromIndex(uint8_t color_index) {
        Color color;
        color.type_ = Type::Indexed;
        color.value_ = color_index;
        color.tint_ = 0.0;
        return color;
    }
    
    /**
     * @brief 自动颜色构造函数
     */
    static Color automatic() {
        Color color;
        color.type_ = Type::Auto;
        color.value_ = 0;
        color.tint_ = 0.0;
        return color;
    }
    
    /**
     * @brief 从十六进制字符串创建颜色
     * @param hex_string 十六进制字符串（如"FF0000"或"#FF0000"）
     */
    static Color fromHex(const std::string& hex_string);
    
    // 预定义颜色常量
    
    static const Color BLACK;
    static const Color WHITE;
    static const Color RED;
    static const Color GREEN;
    static const Color BLUE;
    static const Color YELLOW;
    static const Color MAGENTA;
    static const Color CYAN;
    static const Color BROWN;
    static const Color GRAY;
    static const Color LIME;
    static const Color NAVY;
    static const Color ORANGE;
    static const Color PINK;
    static const Color PURPLE;
    static const Color SILVER;
    
    // 访问器
    
    Type getType() const { return type_; }
    uint32_t getValue() const { return value_; }
    double getTint() const { return tint_; }
    
    /**
     * @brief 获取RGB值
     * @return RGB值（如果不是RGB类型则转换）
     */
    uint32_t getRGB() const;
    
    // 注意：getRGB 基于静态映射，不会引用工作簿主题。
    // 如需按当前主题解析，请使用 ThemeUtils::resolveRGB(color, theme_ptr)。
    
    /**
     * @brief 获取红色分量
     * @return 红色分量 (0-255)
     */
    uint8_t getRed() const { return (getRGB() >> 16) & 0xFF; }
    
    /**
     * @brief 获取绿色分量
     * @return 绿色分量 (0-255)
     */
    uint8_t getGreen() const { return (getRGB() >> 8) & 0xFF; }
    
    /**
     * @brief 获取蓝色分量
     * @return 蓝色分量 (0-255)
     */
    uint8_t getBlue() const { return getRGB() & 0xFF; }
    
    // 修改器
    
    /**
     * @brief 设置色调
     * @param tint 色调值 (-1.0 到 1.0)
     */
    void setTint(double tint) { tint_ = std::max(-1.0, std::min(1.0, tint)); }
    
    /**
     * @brief 调整亮度
     * @param factor 亮度因子 (0.0-2.0, 1.0为原始亮度)
     * @return 新的颜色对象
     */
    Color adjustBrightness(double factor) const;
    
    /**
     * @brief 调整饱和度
     * @param factor 饱和度因子 (0.0-2.0, 1.0为原始饱和度)
     * @return 新的颜色对象
     */
    Color adjustSaturation(double factor) const;
    
    // 格式化
    
    /**
     * @brief 转换为十六进制字符串
     * @param include_hash 是否包含#前缀
     * @return 十六进制字符串
     */
    std::string toHex(bool include_hash = false) const;
    
    /**
     * @brief 转换为Excel XML格式
     * @return XML字符串
     */
    std::string toXML() const;
    
    /**
     * @brief 转换为CSS格式
     * @return CSS颜色字符串
     */
    std::string toCSS() const;
    
    // 比较操作
    
    bool operator==(const Color& other) const {
        return type_ == other.type_ && value_ == other.value_ && tint_ == other.tint_;
    }
    
    bool operator!=(const Color& other) const {
        return !(*this == other);
    }
    
    /**
     * @brief 计算哈希值
     * @return 哈希值
     */
    size_t hash() const {
        return std::hash<uint32_t>{}(static_cast<uint32_t>(type_) | (value_ << 8)) ^ 
               std::hash<double>{}(tint_);
    }
    
    // 类型转换
    
    /**
     * @brief 隐式转换为uint32_t（RGB值）
     */
    operator uint32_t() const { return getRGB(); }
    
    // 颜色混合
    
    /**
     * @brief 与另一个颜色混合
     * @param other 另一个颜色
     * @param ratio 混合比例 (0.0-1.0)
     * @return 混合后的颜色
     */
    Color blend(const Color& other, double ratio) const;
    
    // 颜色空间转换
    
    /**
     * @brief 转换为HSL颜色空间
     * @param h 色相 (0-360)
     * @param s 饱和度 (0-100)
     * @param l 亮度 (0-100)
     */
    void toHSL(double& h, double& s, double& l) const;
    
    /**
     * @brief 从HSL颜色空间创建颜色
     * @param h 色相 (0-360)
     * @param s 饱和度 (0-100)
     * @param l 亮度 (0-100)
     * @return 颜色对象
     */
    static Color fromHSL(double h, double s, double l);

private:
    // 内部辅助方法
    uint32_t themeToRGB() const;
    uint32_t indexedToRGB() const;
    uint32_t applyTint(uint32_t rgb, double tint) const;
    
    // HSL转换辅助函数
    static double hueToRGB(double p, double q, double t);
};

}} // namespace fastexcel::core
