#include "FormatDescriptor.hpp"
#include <memory>

namespace fastexcel {
namespace core {

// 默认值常量
namespace {
    const std::string DEFAULT_FONT_NAME = "Calibri";
    const double DEFAULT_FONT_SIZE = 11.0;
    const core::Color DEFAULT_FONT_COLOR = core::Color::BLACK;
    const core::Color DEFAULT_BG_COLOR = core::Color::WHITE;
    const core::Color DEFAULT_FG_COLOR = core::Color::BLACK;
    const core::Color DEFAULT_BORDER_COLOR = core::Color::BLACK;
}

FormatDescriptor::FormatDescriptor(
    const std::string& font_name,
    double font_size,
    bool bold,
    bool italic,
    UnderlineType underline,
    bool strikeout,
    FontScript script,
    core::Color font_color,
    uint8_t font_family,
    uint8_t font_charset,
    HorizontalAlign horizontal_align,
    VerticalAlign vertical_align,
    bool text_wrap,
    int16_t rotation,
    uint8_t indent,
    bool shrink,
    BorderStyle left_border,
    BorderStyle right_border,
    BorderStyle top_border,
    BorderStyle bottom_border,
    BorderStyle diag_border,
    DiagonalBorderType diag_type,
    core::Color left_border_color,
    core::Color right_border_color,
    core::Color top_border_color,
    core::Color bottom_border_color,
    core::Color diag_border_color,
    PatternType pattern,
    core::Color bg_color,
    core::Color fg_color,
    const std::string& num_format,
    uint16_t num_format_index,
    bool locked,
    bool hidden
) : font_name_(font_name),
    font_size_(font_size),
    bold_(bold),
    italic_(italic),
    underline_(underline),
    strikeout_(strikeout),
    script_(script),
    font_color_(font_color),
    font_family_(font_family),
    font_charset_(font_charset),
    horizontal_align_(horizontal_align),
    vertical_align_(vertical_align),
    text_wrap_(text_wrap),
    rotation_(rotation),
    indent_(indent),
    shrink_(shrink),
    left_border_(left_border),
    right_border_(right_border),
    top_border_(top_border),
    bottom_border_(bottom_border),
    diag_border_(diag_border),
    diag_type_(diag_type),
    left_border_color_(left_border_color),
    right_border_color_(right_border_color),
    top_border_color_(top_border_color),
    bottom_border_color_(bottom_border_color),
    diag_border_color_(diag_border_color),
    pattern_(pattern),
    bg_color_(bg_color),
    fg_color_(fg_color),
    num_format_(num_format),
    num_format_index_(num_format_index),
    locked_(locked),
    hidden_(hidden),
    hash_value_(calculateHash()) {
}

const FormatDescriptor& FormatDescriptor::getDefault() {
    static const FormatDescriptor default_format(
        DEFAULT_FONT_NAME,           // font_name
        DEFAULT_FONT_SIZE,           // font_size
        false,                       // bold
        false,                       // italic
        UnderlineType::None,         // underline
        false,                       // strikeout
        FontScript::None,            // script
        DEFAULT_FONT_COLOR,          // font_color
        2,                           // font_family
        1,                           // font_charset
        HorizontalAlign::None,       // horizontal_align
        VerticalAlign::Bottom,       // vertical_align
        false,                       // text_wrap
        0,                           // rotation
        0,                           // indent
        false,                       // shrink
        BorderStyle::None,           // left_border
        BorderStyle::None,           // right_border
        BorderStyle::None,           // top_border
        BorderStyle::None,           // bottom_border
        BorderStyle::None,           // diag_border
        DiagonalBorderType::None,    // diag_type
        DEFAULT_BORDER_COLOR,        // left_border_color
        DEFAULT_BORDER_COLOR,        // right_border_color
        DEFAULT_BORDER_COLOR,        // top_border_color
        DEFAULT_BORDER_COLOR,        // bottom_border_color
        DEFAULT_BORDER_COLOR,        // diag_border_color
        PatternType::None,           // pattern
        DEFAULT_BG_COLOR,            // bg_color
        DEFAULT_FG_COLOR,            // fg_color
        "",                          // num_format
        0,                           // num_format_index
        true,                        // locked
        false                        // hidden
    );
    return default_format;
}

bool FormatDescriptor::hasFont() const {
    const auto& default_desc = getDefault();
    return font_name_ != default_desc.font_name_ ||
           font_size_ != default_desc.font_size_ ||
           bold_ != default_desc.bold_ ||
           italic_ != default_desc.italic_ ||
           underline_ != default_desc.underline_ ||
           strikeout_ != default_desc.strikeout_ ||
           script_ != default_desc.script_ ||
           font_color_ != default_desc.font_color_;
}

bool FormatDescriptor::hasFill() const {
    const auto& default_desc = getDefault();
    return pattern_ != default_desc.pattern_ ||
           bg_color_ != default_desc.bg_color_ ||
           fg_color_ != default_desc.fg_color_;
}

bool FormatDescriptor::hasBorder() const {
    const auto& default_desc = getDefault();
    return left_border_ != default_desc.left_border_ ||
           right_border_ != default_desc.right_border_ ||
           top_border_ != default_desc.top_border_ ||
           bottom_border_ != default_desc.bottom_border_ ||
           diag_border_ != default_desc.diag_border_ ||
           diag_type_ != default_desc.diag_type_;
}

bool FormatDescriptor::hasAlignment() const {
    const auto& default_desc = getDefault();
    return horizontal_align_ != default_desc.horizontal_align_ ||
           vertical_align_ != default_desc.vertical_align_ ||
           text_wrap_ != default_desc.text_wrap_ ||
           rotation_ != default_desc.rotation_ ||
           indent_ != default_desc.indent_ ||
           shrink_ != default_desc.shrink_;
}

bool FormatDescriptor::hasProtection() const {
    const auto& default_desc = getDefault();
    return locked_ != default_desc.locked_ ||
           hidden_ != default_desc.hidden_;
}

bool FormatDescriptor::hasAnyFormatting() const {
    return hasFont() || hasFill() || hasBorder() || hasAlignment() || hasProtection() || 
           !num_format_.empty() || num_format_index_ != 0;
}

bool FormatDescriptor::operator==(const FormatDescriptor& other) const {
    // 首先比较哈希值，如果不同则快速返回false
    if (hash_value_ != other.hash_value_) {
        return false;
    }
    
    // 如果哈希值相同，进行完整比较（处理哈希冲突）
    return font_name_ == other.font_name_ &&
           font_size_ == other.font_size_ &&
           bold_ == other.bold_ &&
           italic_ == other.italic_ &&
           underline_ == other.underline_ &&
           strikeout_ == other.strikeout_ &&
           script_ == other.script_ &&
           font_color_ == other.font_color_ &&
           font_family_ == other.font_family_ &&
           font_charset_ == other.font_charset_ &&
           horizontal_align_ == other.horizontal_align_ &&
           vertical_align_ == other.vertical_align_ &&
           text_wrap_ == other.text_wrap_ &&
           rotation_ == other.rotation_ &&
           indent_ == other.indent_ &&
           shrink_ == other.shrink_ &&
           left_border_ == other.left_border_ &&
           right_border_ == other.right_border_ &&
           top_border_ == other.top_border_ &&
           bottom_border_ == other.bottom_border_ &&
           diag_border_ == other.diag_border_ &&
           diag_type_ == other.diag_type_ &&
           left_border_color_ == other.left_border_color_ &&
           right_border_color_ == other.right_border_color_ &&
           top_border_color_ == other.top_border_color_ &&
           bottom_border_color_ == other.bottom_border_color_ &&
           diag_border_color_ == other.diag_border_color_ &&
           pattern_ == other.pattern_ &&
           bg_color_ == other.bg_color_ &&
           fg_color_ == other.fg_color_ &&
           num_format_ == other.num_format_ &&
           num_format_index_ == other.num_format_index_ &&
           locked_ == other.locked_ &&
           hidden_ == other.hidden_;
}

size_t FormatDescriptor::calculateHash() const {
    std::hash<std::string> str_hasher;
    std::hash<double> double_hasher;
    std::hash<bool> bool_hasher;
    std::hash<uint8_t> uint8_hasher;
    std::hash<uint16_t> uint16_hasher;
    std::hash<int16_t> int16_hasher;
    std::hash<size_t> size_hasher;
    
    size_t h1 = str_hasher(font_name_);
    size_t h2 = double_hasher(font_size_);
    size_t h3 = bool_hasher(bold_);
    size_t h4 = bool_hasher(italic_);
    size_t h5 = static_cast<size_t>(underline_);
    size_t h6 = bool_hasher(strikeout_);
    size_t h7 = static_cast<size_t>(script_);
    size_t h8 = font_color_.hash();
    size_t h9 = uint8_hasher(font_family_);
    size_t h10 = uint8_hasher(font_charset_);
    
    size_t h11 = static_cast<size_t>(horizontal_align_);
    size_t h12 = static_cast<size_t>(vertical_align_);
    size_t h13 = bool_hasher(text_wrap_);
    size_t h14 = int16_hasher(rotation_);
    size_t h15 = uint8_hasher(indent_);
    size_t h16 = bool_hasher(shrink_);
    
    size_t h17 = static_cast<size_t>(left_border_);
    size_t h18 = static_cast<size_t>(right_border_);
    size_t h19 = static_cast<size_t>(top_border_);
    size_t h20 = static_cast<size_t>(bottom_border_);
    size_t h21 = static_cast<size_t>(diag_border_);
    size_t h22 = static_cast<size_t>(diag_type_);
    
    size_t h23 = left_border_color_.hash();
    size_t h24 = right_border_color_.hash();
    size_t h25 = top_border_color_.hash();
    size_t h26 = bottom_border_color_.hash();
    size_t h27 = diag_border_color_.hash();
    
    size_t h28 = static_cast<size_t>(pattern_);
    size_t h29 = bg_color_.hash();
    size_t h30 = fg_color_.hash();
    
    size_t h31 = str_hasher(num_format_);
    size_t h32 = uint16_hasher(num_format_index_);
    
    size_t h33 = bool_hasher(locked_);
    size_t h34 = bool_hasher(hidden_);
    
    // 使用改进的哈希组合算法
    size_t result = 0;
    size_t hashes[] = {h1, h2, h3, h4, h5, h6, h7, h8, h9, h10,
                       h11, h12, h13, h14, h15, h16, h17, h18, h19, h20,
                       h21, h22, h23, h24, h25, h26, h27, h28, h29, h30,
                       h31, h32, h33, h34};
    
    for (size_t i = 0; i < sizeof(hashes) / sizeof(hashes[0]); ++i) {
        result ^= hashes[i] + 0x9e3779b9 + (result << 6) + (result >> 2);
    }
    
    return result;
}

}} // namespace fastexcel::core