#include "fastexcel/core/Color.hpp"
#include <sstream>
#include <iomanip>
#include <algorithm>
#include <cmath>

namespace fastexcel {
namespace core {

// ========== 预定义颜色常量 ==========

const Color Color::BLACK(static_cast<uint32_t>(0x000000));
const Color Color::WHITE(static_cast<uint32_t>(0xFFFFFF));
const Color Color::RED(static_cast<uint32_t>(0xFF0000));
const Color Color::GREEN(static_cast<uint32_t>(0x008000));
const Color Color::BLUE(static_cast<uint32_t>(0x0000FF));
const Color Color::YELLOW(static_cast<uint32_t>(0xFFFF00));
const Color Color::MAGENTA(static_cast<uint32_t>(0xFF00FF));
const Color Color::CYAN(static_cast<uint32_t>(0x00FFFF));
const Color Color::BROWN(static_cast<uint32_t>(0x800000));
const Color Color::GRAY(static_cast<uint32_t>(0x808080));
const Color Color::LIME(static_cast<uint32_t>(0x00FF00));
const Color Color::NAVY(static_cast<uint32_t>(0x000080));
const Color Color::ORANGE(static_cast<uint32_t>(0xFF6600));
const Color Color::PINK(static_cast<uint32_t>(0xFF00FF));
const Color Color::PURPLE(static_cast<uint32_t>(0x800080));
const Color Color::SILVER(static_cast<uint32_t>(0xC0C0C0));

// ========== 静态方法 ==========

Color Color::fromHex(const std::string& hex_string) {
    std::string hex = hex_string;
    
    // 移除#前缀
    if (!hex.empty() && hex[0] == '#') {
        hex = hex.substr(1);
    }
    
    // 确保长度为6
    if (hex.length() != 6) {
        return Color::BLACK;  // 默认返回黑色
    }
    
    // 转换为数值
    uint32_t rgb = 0;
    std::stringstream ss;
    ss << std::hex << hex;
    ss >> rgb;
    
    return Color(rgb);
}

Color Color::fromHSL(double h, double s, double l) {
    // 将HSL值标准化
    h = fmod(h, 360.0);
    if (h < 0) h += 360.0;
    s = std::max(0.0, std::min(100.0, s)) / 100.0;
    l = std::max(0.0, std::min(100.0, l)) / 100.0;
    
    double r, g, b;
    
    if (s == 0.0) {
        // 灰度
        r = g = b = l;
    } else {
        double q = l < 0.5 ? l * (1.0 + s) : l + s - l * s;
        double p = 2.0 * l - q;
        
        double h_norm = h / 360.0;
        r = hueToRGB(p, q, h_norm + 1.0/3.0);
        g = hueToRGB(p, q, h_norm);
        b = hueToRGB(p, q, h_norm - 1.0/3.0);
    }
    
    // 转换为RGB值
    uint32_t rgb = (static_cast<uint32_t>(r * 255) << 16) |
                   (static_cast<uint32_t>(g * 255) << 8) |
                   static_cast<uint32_t>(b * 255);
    
    return Color(rgb);
}

// ========== 访问器 ==========

uint32_t Color::getRGB() const {
    switch (type_) {
        case Type::RGB:
            return applyTint(value_, tint_);
        case Type::Theme:
            return applyTint(themeToRGB(), tint_);
        case Type::Indexed:
            return applyTint(indexedToRGB(), tint_);
        case Type::Auto:
            return 0x000000;  // 自动颜色默认为黑色
        default:
            return 0x000000;
    }
}

// ========== 修改器 ==========

Color Color::adjustBrightness(double factor) const {
    uint32_t rgb = getRGB();
    uint8_t r = (rgb >> 16) & 0xFF;
    uint8_t g = (rgb >> 8) & 0xFF;
    uint8_t b = rgb & 0xFF;
    
    r = static_cast<uint8_t>(std::min(255.0, r * factor));
    g = static_cast<uint8_t>(std::min(255.0, g * factor));
    b = static_cast<uint8_t>(std::min(255.0, b * factor));
    
    return Color(static_cast<uint32_t>((r << 16) | (g << 8) | b));
}

Color Color::adjustSaturation(double factor) const {
    double h, s, l;
    toHSL(h, s, l);
    s = std::min(100.0, s * factor);
    return fromHSL(h, s, l);
}

// ========== 格式化 ==========

std::string Color::toHex(bool include_hash) const {
    uint32_t rgb = getRGB();
    std::ostringstream oss;
    if (include_hash) {
        oss << "#";
    }
    oss << std::hex << std::uppercase << std::setw(6) << std::setfill('0') << rgb;
    return oss.str();
}

std::string Color::toXML() const {
    std::ostringstream oss;
    
    switch (type_) {
        case Type::RGB:
            oss << "<color rgb=\"" << toHex() << "\"/>";
            break;
        case Type::Theme:
            oss << "<color theme=\"" << value_ << "\"";
            if (tint_ != 0.0) {
                oss << " tint=\"" << tint_ << "\"";
            }
            oss << "/>";
            break;
        case Type::Indexed:
            oss << "<color indexed=\"" << value_ << "\"/>";
            break;
        case Type::Auto:
            oss << "<color auto=\"1\"/>";
            break;
    }
    
    return oss.str();
}

std::string Color::toCSS() const {
    uint32_t rgb = getRGB();
    uint8_t r = (rgb >> 16) & 0xFF;
    uint8_t g = (rgb >> 8) & 0xFF;
    uint8_t b = rgb & 0xFF;
    
    std::ostringstream oss;
    oss << "rgb(" << static_cast<int>(r) << ", " 
        << static_cast<int>(g) << ", " 
        << static_cast<int>(b) << ")";
    return oss.str();
}

// ========== 颜色混合 ==========

Color Color::blend(const Color& other, double ratio) const {
    ratio = std::max(0.0, std::min(1.0, ratio));
    
    uint32_t rgb1 = getRGB();
    uint32_t rgb2 = other.getRGB();
    
    uint8_t r1 = (rgb1 >> 16) & 0xFF;
    uint8_t g1 = (rgb1 >> 8) & 0xFF;
    uint8_t b1 = rgb1 & 0xFF;
    
    uint8_t r2 = (rgb2 >> 16) & 0xFF;
    uint8_t g2 = (rgb2 >> 8) & 0xFF;
    uint8_t b2 = rgb2 & 0xFF;
    
    uint8_t r = static_cast<uint8_t>(r1 * (1.0 - ratio) + r2 * ratio);
    uint8_t g = static_cast<uint8_t>(g1 * (1.0 - ratio) + g2 * ratio);
    uint8_t b = static_cast<uint8_t>(b1 * (1.0 - ratio) + b2 * ratio);
    
    return Color(static_cast<uint32_t>((r << 16) | (g << 8) | b));
}

// ========== 颜色空间转换 ==========

void Color::toHSL(double& h, double& s, double& l) const {
    uint32_t rgb = getRGB();
    double r = ((rgb >> 16) & 0xFF) / 255.0;
    double g = ((rgb >> 8) & 0xFF) / 255.0;
    double b = (rgb & 0xFF) / 255.0;
    
    double max_val = std::max({r, g, b});
    double min_val = std::min({r, g, b});
    double delta = max_val - min_val;
    
    // 亮度
    l = (max_val + min_val) / 2.0 * 100.0;
    
    if (delta == 0.0) {
        // 灰度
        h = s = 0.0;
    } else {
        // 饱和度
        s = (l > 50.0) ? delta / (2.0 - max_val - min_val) : delta / (max_val + min_val);
        s *= 100.0;
        
        // 色相
        if (max_val == r) {
            h = ((g - b) / delta) + (g < b ? 6.0 : 0.0);
        } else if (max_val == g) {
            h = (b - r) / delta + 2.0;
        } else {
            h = (r - g) / delta + 4.0;
        }
        h *= 60.0;
    }
}

// ========== 私有方法 ==========

uint32_t Color::themeToRGB() const {
    // Excel主题颜色映射表（简化版）
    static const uint32_t theme_colors[] = {
        0x000000, // 0: 深色1
        0xFFFFFF, // 1: 浅色1
        0x1F497D, // 2: 深色2
        0xEEECE1, // 3: 浅色2
        0x4F81BD, // 4: 强调文字颜色1
        0xF79646, // 5: 强调文字颜色2
        0x9BBB59, // 6: 强调文字颜色3
        0x8064A2, // 7: 强调文字颜色4
        0x4BACC6, // 8: 强调文字颜色5
        0xF366A7, // 9: 强调文字颜色6
        0x0000FF, // 10: 超链接
        0x800080  // 11: 已访问的超链接
    };
    
    if (value_ < sizeof(theme_colors) / sizeof(theme_colors[0])) {
        return theme_colors[value_];
    }
    return 0x000000;  // 默认黑色
}

uint32_t Color::indexedToRGB() const {
    // Excel索引颜色映射表（简化版）
    static const uint32_t indexed_colors[] = {
        0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF,
        0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF,
        0x800000, 0x008000, 0x000080, 0x808000, 0x800080, 0x008080, 0xC0C0C0, 0x808080,
        0x9999FF, 0x993366, 0xFFFFCC, 0xCCFFFF, 0x660066, 0xFF8080, 0x0066CC, 0xCCCCFF,
        0x000080, 0xFF00FF, 0xFFFF00, 0x00FFFF, 0x800080, 0x800000, 0x008080, 0x0000FF,
        0x00CCFF, 0xCCFFFF, 0xCCFFCC, 0xFFFF99, 0x99CCFF, 0xFF99CC, 0xCC99FF, 0xFFCC99,
        0x3366FF, 0x33CCCC, 0x99CC00, 0xFFCC00, 0xFF9900, 0xFF6600, 0x666699, 0x969696,
        0x003366, 0x339966, 0x003300, 0x333300, 0x993300, 0x993366, 0x333399, 0x333333
    };
    
    if (value_ < sizeof(indexed_colors) / sizeof(indexed_colors[0])) {
        return indexed_colors[value_];
    }
    return 0x000000;  // 默认黑色
}

uint32_t Color::applyTint(uint32_t rgb, double tint) const {
    if (tint == 0.0) {
        return rgb;
    }
    
    uint8_t r = (rgb >> 16) & 0xFF;
    uint8_t g = (rgb >> 8) & 0xFF;
    uint8_t b = rgb & 0xFF;
    
    if (tint > 0.0) {
        // 变亮
        r = static_cast<uint8_t>(r + (255 - r) * tint);
        g = static_cast<uint8_t>(g + (255 - g) * tint);
        b = static_cast<uint8_t>(b + (255 - b) * tint);
    } else {
        // 变暗
        double factor = 1.0 + tint;
        r = static_cast<uint8_t>(r * factor);
        g = static_cast<uint8_t>(g * factor);
        b = static_cast<uint8_t>(b * factor);
    }
    
    return static_cast<uint32_t>((r << 16) | (g << 8) | b);
}

double Color::hueToRGB(double p, double q, double t) {
    if (t < 0.0) t += 1.0;
    if (t > 1.0) t -= 1.0;
    if (t < 1.0/6.0) return p + (q - p) * 6.0 * t;
    if (t < 1.0/2.0) return q;
    if (t < 2.0/3.0) return p + (q - p) * (2.0/3.0 - t) * 6.0;
    return p;
}

}} // namespace fastexcel::core