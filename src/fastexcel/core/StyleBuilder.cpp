#include "StyleBuilder.hpp"

namespace fastexcel {
namespace core {

StyleBuilder::StyleBuilder(const domain::FormatDescriptor& format) 
    : font_name_(format.getFontName()),
      font_size_(format.getFontSize()),
      bold_(format.isBold()),
      italic_(format.isItalic()),
      underline_(format.getUnderline()),
      strikeout_(format.isStrikeout()),
      script_(format.getFontScript()),
      font_color_(format.getFontColor()),
      font_family_(format.getFontFamily()),
      font_charset_(format.getFontCharset()),
      horizontal_align_(format.getHorizontalAlign()),
      vertical_align_(format.getVerticalAlign()),
      text_wrap_(format.isTextWrap()),
      rotation_(format.getRotation()),
      indent_(format.getIndent()),
      shrink_(format.isShrink()),
      left_border_(format.getLeftBorder()),
      right_border_(format.getRightBorder()),
      top_border_(format.getTopBorder()),
      bottom_border_(format.getBottomBorder()),
      diag_border_(format.getDiagBorder()),
      diag_type_(format.getDiagType()),
      left_border_color_(format.getLeftBorderColor()),
      right_border_color_(format.getRightBorderColor()),
      top_border_color_(format.getTopBorderColor()),
      bottom_border_color_(format.getBottomBorderColor()),
      diag_border_color_(format.getDiagBorderColor()),
      pattern_(format.getPattern()),
      bg_color_(format.getBackgroundColor()),
      fg_color_(format.getForegroundColor()),
      num_format_(format.getNumberFormat()),
      num_format_index_(format.getNumberFormatIndex()),
      locked_(format.isLocked()),
      hidden_(format.isHidden()) {
}

domain::FormatDescriptor StyleBuilder::build() const {
    // 这里我们直接调用FormatDescriptor的私有构造函数
    // 由于StyleBuilder是FormatDescriptor的友元类，这是允许的
    return domain::FormatDescriptor(
        font_name_,
        font_size_,
        bold_,
        italic_,
        underline_,
        strikeout_,
        script_,
        font_color_,
        font_family_,
        font_charset_,
        horizontal_align_,
        vertical_align_,
        text_wrap_,
        rotation_,
        indent_,
        shrink_,
        left_border_,
        right_border_,
        top_border_,
        bottom_border_,
        diag_border_,
        diag_type_,
        left_border_color_,
        right_border_color_,
        top_border_color_,
        bottom_border_color_,
        diag_border_color_,
        pattern_,
        bg_color_,
        fg_color_,
        num_format_,
        num_format_index_,
        locked_,
        hidden_
    );
}

}} // namespace fastexcel::core