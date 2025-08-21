#include "FormatTypes.hpp"

namespace fastexcel {
namespace core {

// toString 函数实现

const char* toString(UnderlineType type) {
    switch (type) {
        case UnderlineType::None: return "None";
        case UnderlineType::Single: return "Single";
        case UnderlineType::Double: return "Double";
        case UnderlineType::SingleAccounting: return "SingleAccounting";
        case UnderlineType::DoubleAccounting: return "DoubleAccounting";
        default: return "Unknown";
    }
}

const char* toString(FontScript script) {
    switch (script) {
        case FontScript::None: return "None";
        case FontScript::Superscript: return "Superscript";
        case FontScript::Subscript: return "Subscript";
        default: return "Unknown";
    }
}

const char* toString(HorizontalAlign align) {
    switch (align) {
        case HorizontalAlign::None: return "None";
        case HorizontalAlign::Left: return "Left";
        case HorizontalAlign::Center: return "Center";
        case HorizontalAlign::Right: return "Right";
        case HorizontalAlign::Fill: return "Fill";
        case HorizontalAlign::Justify: return "Justify";
        case HorizontalAlign::CenterAcross: return "CenterAcross";
        case HorizontalAlign::Distributed: return "Distributed";
        default: return "Unknown";
    }
}

const char* toString(VerticalAlign align) {
    switch (align) {
        case VerticalAlign::Top: return "Top";
        case VerticalAlign::Center: return "Center";
        case VerticalAlign::Bottom: return "Bottom";
        case VerticalAlign::Justify: return "Justify";
        case VerticalAlign::Distributed: return "Distributed";
        default: return "Unknown";
    }
}

const char* toString(PatternType pattern) {
    switch (pattern) {
        case PatternType::None: return "None";
        case PatternType::Solid: return "Solid";
        case PatternType::MediumGray: return "MediumGray";
        case PatternType::DarkGray: return "DarkGray";
        case PatternType::LightGray: return "LightGray";
        case PatternType::DarkHorizontal: return "DarkHorizontal";
        case PatternType::DarkVertical: return "DarkVertical";
        case PatternType::DarkDown: return "DarkDown";
        case PatternType::DarkUp: return "DarkUp";
        case PatternType::DarkGrid: return "DarkGrid";
        case PatternType::DarkTrellis: return "DarkTrellis";
        case PatternType::LightHorizontal: return "LightHorizontal";
        case PatternType::LightVertical: return "LightVertical";
        case PatternType::LightDown: return "LightDown";
        case PatternType::LightUp: return "LightUp";
        case PatternType::LightGrid: return "LightGrid";
        case PatternType::LightTrellis: return "LightTrellis";
        case PatternType::Gray125: return "Gray125";
        case PatternType::Gray0625: return "Gray0625";
        default: return "Unknown";
    }
}

const char* toString(BorderStyle style) {
    switch (style) {
        case BorderStyle::None: return "None";
        case BorderStyle::Thin: return "Thin";
        case BorderStyle::Medium: return "Medium";
        case BorderStyle::Thick: return "Thick";
        case BorderStyle::Double: return "Double";
        case BorderStyle::Hair: return "Hair";
        case BorderStyle::Dotted: return "Dotted";
        case BorderStyle::Dashed: return "Dashed";
        case BorderStyle::DashDot: return "DashDot";
        case BorderStyle::DashDotDot: return "DashDotDot";
        case BorderStyle::MediumDashed: return "MediumDashed";
        case BorderStyle::MediumDashDot: return "MediumDashDot";
        case BorderStyle::MediumDashDotDot: return "MediumDashDotDot";
        case BorderStyle::SlantDashDot: return "SlantDashDot";
        default: return "Unknown";
    }
}

const char* toString(DiagonalBorderType type) {
    switch (type) {
        case DiagonalBorderType::None: return "None";
        case DiagonalBorderType::Up: return "Up";
        case DiagonalBorderType::Down: return "Down";
        case DiagonalBorderType::Both: return "Both";
        default: return "Unknown";
    }
}

const char* toString(NumberFormatType type) {
    switch (type) {
        case NumberFormatType::General: return "General";
        case NumberFormatType::Number: return "Number";
        case NumberFormatType::Decimal: return "Decimal";
        case NumberFormatType::Currency: return "Currency";
        case NumberFormatType::Accounting: return "Accounting";
        case NumberFormatType::Date: return "Date";
        case NumberFormatType::Time: return "Time";
        case NumberFormatType::Percentage: return "Percentage";
        case NumberFormatType::Fraction: return "Fraction";
        case NumberFormatType::Scientific: return "Scientific";
        case NumberFormatType::Text: return "Text";
        default: return "Unknown";
    }
}

// isValid 函数实现

bool isValid(UnderlineType type) {
    return static_cast<uint8_t>(type) <= static_cast<uint8_t>(UnderlineType::DoubleAccounting);
}

bool isValid(FontScript script) {
    return static_cast<uint8_t>(script) <= static_cast<uint8_t>(FontScript::Subscript);
}

bool isValid(HorizontalAlign align) {
    return static_cast<uint8_t>(align) <= static_cast<uint8_t>(HorizontalAlign::Distributed);
}

bool isValid(VerticalAlign align) {
    return static_cast<uint8_t>(align) <= static_cast<uint8_t>(VerticalAlign::Distributed);
}

bool isValid(PatternType pattern) {
    return static_cast<uint8_t>(pattern) <= static_cast<uint8_t>(PatternType::Gray0625);
}

bool isValid(BorderStyle style) {
    return static_cast<uint8_t>(style) <= static_cast<uint8_t>(BorderStyle::SlantDashDot);
}

bool isValid(DiagonalBorderType type) {
    return static_cast<uint8_t>(type) <= static_cast<uint8_t>(DiagonalBorderType::Both);
}

bool isValid(NumberFormatType type) {
    // NumberFormatType 使用Excel内置格式编号，需要特殊验证
    switch (type) {
        case NumberFormatType::General:
        case NumberFormatType::Number:
        case NumberFormatType::Decimal:
        case NumberFormatType::Currency:
        case NumberFormatType::Accounting:
        case NumberFormatType::Date:
        case NumberFormatType::Time:
        case NumberFormatType::Percentage:
        case NumberFormatType::Fraction:
        case NumberFormatType::Scientific:
        case NumberFormatType::Text:
            return true;
        default:
            return false;
    }
}

}} // namespace fastexcel::core
