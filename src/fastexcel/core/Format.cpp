#include "fastexcel/core/Format.hpp"
#include <sstream>
#include <iomanip>

namespace fastexcel {
namespace core {

std::string Format::generateFontXML() const {
    std::ostringstream oss;
    
    oss << "<font>";
    
    if (!font_name_.empty()) {
        oss << "<name val=\"" << font_name_ << "\"/>";
    }
    
    if (font_size_ > 0) {
        oss << "<sz val=\"" << font_size_ << "\"/>";
    }
    
    if (font_bold_ == FontBold::Bold) {
        oss << "<b/>";
    }
    
    if (font_italic_ == FontItalic::Italic) {
        oss << "<i/>";
    }
    
    if (font_underline_ != FontUnderline::None) {
        oss << "<u val=\"" << (font_underline_ == FontUnderline::Single ? "single" : "double") << "\"/>";
    }
    
    if (font_color_ != 0x000000) {
        oss << "<color rgb=\"" << std::hex << std::setw(6) << std::setfill('0') << font_color_ << "\"/>";
    }
    
    oss << "</font>";
    
    return oss.str();
}

std::string Format::generateAlignmentXML() const {
    std::ostringstream oss;
    
    bool has_alignment = (horizontal_align_ != HorizontalAlignment::General ||
                         vertical_align_ != VerticalAlignment::Bottom ||
                         wrap_text_);
    
    if (!has_alignment) {
        return "";
    }
    
    oss << "<alignment";
    
    if (horizontal_align_ != HorizontalAlignment::General) {
        std::string align_str;
        switch (horizontal_align_) {
            case HorizontalAlignment::Left: align_str = "left"; break;
            case HorizontalAlignment::Center: align_str = "center"; break;
            case HorizontalAlignment::Right: align_str = "right"; break;
            case HorizontalAlignment::Fill: align_str = "fill"; break;
            case HorizontalAlignment::Justify: align_str = "justify"; break;
            case HorizontalAlignment::CenterContinuous: align_str = "centerContinuous"; break;
            case HorizontalAlignment::Distributed: align_str = "distributed"; break;
            default: align_str = "general"; break;
        }
        oss << " horizontal=\"" << align_str << "\"";
    }
    
    if (vertical_align_ != VerticalAlignment::Bottom) {
        std::string align_str;
        switch (vertical_align_) {
            case VerticalAlignment::Top: align_str = "top"; break;
            case VerticalAlignment::Center: align_str = "center"; break;
            case VerticalAlignment::Justify: align_str = "justify"; break;
            case VerticalAlignment::Distributed: align_str = "distributed"; break;
            default: align_str = "bottom"; break;
        }
        oss << " vertical=\"" << align_str << "\"";
    }
    
    if (wrap_text_) {
        oss << " wrapText=\"1\"";
    }
    
    oss << "/>";
    
    return oss.str();
}

std::string Format::generateFillXML() const {
    std::ostringstream oss;
    
    if (!has_background_) {
        return "";
    }
    
    oss << "<fill>";
    oss << "<patternFill patternType=\"" << pattern_type_ << "\">";
    
    if (pattern_type_ != "none") {
        if (background_color_ != 0xFFFFFF) {
            oss << "<fgColor rgb=\"" << std::hex << std::setw(6) << std::setfill('0') << background_color_ << "\"/>";
        }
        if (pattern_color_ != 0x000000) {
            oss << "<bgColor rgb=\"" << std::hex << std::setw(6) << std::setfill('0') << pattern_color_ << "\"/>";
        }
    }
    
    oss << "</patternFill>";
    oss << "</fill>";
    
    return oss.str();
}

std::string Format::generateBorderXML() const {
    std::ostringstream oss;
    
    if (!has_border_) {
        return "";
    }
    
    oss << "<border>";
    
    // 简化版本，只设置一个统一的边框样式
    if (border_style_ != "none") {
        oss << "<left style=\"" << border_style_ << "\">";
        oss << "<color rgb=\"" << std::hex << std::setw(6) << std::setfill('0') << border_color_ << "\"/>";
        oss << "</left>";
        
        oss << "<right style=\"" << border_style_ << "\">";
        oss << "<color rgb=\"" << std::hex << std::setw(6) << std::setfill('0') << border_color_ << "\"/>";
        oss << "</right>";
        
        oss << "<top style=\"" << border_style_ << "\">";
        oss << "<color rgb=\"" << std::hex << std::setw(6) << std::setfill('0') << border_color_ << "\"/>";
        oss << "</top>";
        
        oss << "<bottom style=\"" << border_style_ << "\">";
        oss << "<color rgb=\"" << std::hex << std::setw(6) << std::setfill('0') << border_color_ << "\"/>";
        oss << "</bottom>";
    } else {
        oss << "<left/><right/><top/><bottom/>";
    }
    
    oss << "<diagonal/>";
    oss << "</border>";
    
    return oss.str();
}

std::string Format::generateNumberFormatXML() const {
    if (number_format_.empty()) {
        return "";
    }
    
    std::ostringstream oss;
    oss << "<numFmt formatCode=\"" << number_format_ << "\" numFmtId=\"" << format_id_ << "\"/>";
    return oss.str();
}

bool Format::hasAnyFormatting() const {
    return !font_name_.empty() ||
           font_size_ != 11 ||
           font_bold_ != FontBold::None ||
           font_italic_ != FontItalic::None ||
           font_underline_ != FontUnderline::None ||
           font_color_ != 0x000000 ||
           horizontal_align_ != HorizontalAlignment::General ||
           vertical_align_ != VerticalAlignment::Bottom ||
           wrap_text_ ||
           has_background_ ||
           has_border_ ||
           !number_format_.empty();
}

}} // namespace fastexcel::core