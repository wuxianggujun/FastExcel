#include "fastexcel/utils/Logger.hpp"
//
// Created by wuxianggujun on 25-8-4.
//

#pragma once

#include "BaseSAXParser.hpp"
#include "fastexcel/core/FormatDescriptor.hpp"
#include "fastexcel/core/Color.hpp"
#include <string>
#include <vector>
#include <unordered_map>
#include <memory>

namespace fastexcel {
namespace reader {

/**
 * @brief 高性能样式解析器 - 基于SAX流式解析
 * 
 * 负责解析Excel文件中的styles.xml文件，
 * 使用BaseSAXParser提供的高性能SAX解析能力，
 * 完全消除字符串查找操作，大幅提升样式解析性能。
 * 
 * 性能优化：
 * - 零字符串查找：基于SAX事件驱动
 * - 流式解析：不需要多次find/substr操作
 * - 状态机驱动：智能处理嵌套XML结构
 * - 内存效率：最小化临时对象创建
 * 
 * 支持完整的Excel样式：
 * - 数字格式 (numFmts)
 * - 字体样式 (fonts) 
 * - 填充样式 (fills)
 * - 边框样式 (borders)
 * - 单元格格式 (cellXfs)
 */
class StylesParser : public BaseSAXParser {
public:
    StylesParser() = default;
    ~StylesParser() = default;
    
    /**
     * @brief 解析样式XML内容
     * @param xml_content XML内容
     * @return 是否解析成功
     */
    bool parse(const std::string& xml_content) {
        clear();
        return parseXML(xml_content);
    }
    
    /**
     * @brief 根据XF索引获取格式描述符
     * @param xf_index XF索引
     * @return 格式描述符指针，如果索引无效返回nullptr
     */
    std::shared_ptr<core::FormatDescriptor> getFormat(int xf_index) const;
    
    /**
     * @brief 获取解析的格式数量
     * @return 格式数量
     */
    size_t getFormatCount() const { return cell_xfs_.size(); }

private:
    // 前向声明
    struct FontInfo;
    struct FillInfo;
    struct BorderInfo;
    struct BorderSide;
    struct CellXf;

    // 解析状态机 - 管理复杂的嵌套XML结构
    struct ParseState {
        enum class Context {
            None,
            NumFmts,        // <numFmts>
            Fonts,          // <fonts>
            Fills,          // <fills> 
            Borders,        // <borders>
            CellXfs,        // <cellXfs>
            
            // 字体子上下文
            Font,           // <font>
            FontName,       // <name>
            FontSize,       // <sz>
            FontColor,      // <color>
            
            // 填充子上下文  
            Fill,           // <fill>
            PatternFill,    // <patternFill>
            FgColor,        // <fgColor>
            BgColor,        // <bgColor>
            
            // 边框子上下文
            Border,         // <border>
            BorderLeft,     // <left>
            BorderRight,    // <right>
            BorderTop,      // <top>
            BorderBottom,   // <bottom>
            BorderDiagonal, // <diagonal>
            BorderColor     // 边框颜色
        };
        
        std::vector<Context> context_stack;
        
        // 当前正在构建的对象
        FontInfo* current_font = nullptr;
        FillInfo* current_fill = nullptr; 
        BorderInfo* current_border = nullptr;
        CellXf* current_xf = nullptr;
        
        // 当前边框边指针
        BorderSide* current_border_side = nullptr;
        
        void reset() {
            context_stack.clear();
            current_font = nullptr;
            current_fill = nullptr;
            current_border = nullptr;
            current_xf = nullptr;
            current_border_side = nullptr;
        }
        
        Context getCurrentContext() const {
            return context_stack.empty() ? Context::None : context_stack.back();
        }
        
        void pushContext(Context ctx) {
            context_stack.push_back(ctx);
        }
        
        void popContext() {
            if (!context_stack.empty()) {
                context_stack.pop_back();
            }
        }
    } parse_state_;
    
    // 字体信息结构
    struct FontInfo {
        std::string name = "Calibri";
        double size = 11.0;
        bool bold = false;
        bool italic = false;
        bool underline = false;
        bool strikeout = false;
        core::Color color;
    };
    
    // 填充信息结构
    struct FillInfo {
        std::string pattern_type = "none";
        core::Color fg_color;
        core::Color bg_color;
    };
    
    // 边框边信息结构
    struct BorderSide {
        std::string style;
        core::Color color;
    };
    
    // 边框信息结构
    struct BorderInfo {
        BorderSide left;
        BorderSide right;
        BorderSide top;
        BorderSide bottom;
        BorderSide diagonal;
    };
    
    // 单元格XF信息结构
    struct CellXf {
        int num_fmt_id = -1;
        int font_id = -1;
        int fill_id = -1;
        int border_id = -1;
        std::string horizontal_alignment;
        std::string vertical_alignment;
        bool wrap_text = false;
        int indent = 0;
        int text_rotation = 0;
    };
    
    // 解析后的数据
    std::vector<FontInfo> fonts_;
    std::vector<FillInfo> fills_;
    std::vector<BorderInfo> borders_;
    std::vector<CellXf> cell_xfs_;
    std::unordered_map<int, std::string> number_formats_;
    
    // 重写基类虚函数
    void onStartElement(const std::string& name, const std::vector<xml::XMLAttribute>& attributes, int depth) override;
    void onEndElement(const std::string& name, int depth) override;
    void onText(const std::string& text, int depth) override;
    
    // 清理数据
    void clear() {
        fonts_.clear();
        fills_.clear();
        borders_.clear();
        cell_xfs_.clear();
        number_formats_.clear();
        parse_state_.reset();
    }
    
    // 枚举转换方法
    core::HorizontalAlign getAlignment(const std::string& alignment) const;
    core::VerticalAlign getVerticalAlignment(const std::string& alignment) const;
    core::BorderStyle getBorderStyle(const std::string& style) const;
    core::PatternType getPatternType(const std::string& pattern) const;
    
    // 内置数字格式
    std::string getBuiltinNumberFormat(int format_id) const;
};

} // namespace reader
} // namespace fastexcel

