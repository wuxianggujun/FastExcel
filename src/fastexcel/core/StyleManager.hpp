#pragma once

#include "fastexcel/core/StyleTemplate.hpp"
#include "fastexcel/core/FormatRepository.hpp"
#include <memory>
#include <unordered_map>
#include <vector>

namespace fastexcel {
namespace core {

/**
 * @brief 高级样式管理器 - 整合样式模板和格式仓储功能
 */
class StyleManager {
private:
    std::unique_ptr<StyleTemplate> template_;
    std::unique_ptr<FormatRepository> format_repo_;
    
    // 样式索引映射：原始索引 -> 新索引
    std::unordered_map<int, size_t> style_index_mapping_;
    
    // 快速样式缓存
    std::unordered_map<std::string, size_t> style_cache_;
    
public:
    StyleManager();
    ~StyleManager() = default;
    
    // 禁用拷贝，允许移动
    StyleManager(const StyleManager&) = delete;
    StyleManager& operator=(const StyleManager&) = delete;
    StyleManager(StyleManager&&) = default;
    StyleManager& operator=(StyleManager&&) = default;
    
    /**
     * @brief 初始化预定义样式
     */
    void initializePredefinedStyles();
    
    /**
     * @brief 从工作簿导入样式
     * @param styles 原始样式数据
     */
    void importStylesFromWorkbook(const std::unordered_map<int, std::shared_ptr<Format>>& styles);
    
    /**
     * @brief 获取样式索引（通过原始索引）
     * @param original_index 原始样式索引
     * @return 新的格式仓储索引
     */
    size_t getStyleIndex(int original_index) const;
    
    /**
     * @brief 获取预定义样式索引
     * @param style_name 样式名称
     * @return 格式仓储索引
     */
    size_t getPredefinedStyleIndex(const std::string& style_name);
    
    /**
     * @brief 创建并缓存字体样式
     * @param style_key 样式键值（用于缓存）
     * @param font_name 字体名称
     * @param font_size 字体大小
     * @param bold 粗体
     * @param italic 斜体
     * @param color 颜色
     * @return 格式仓储索引
     */
    size_t createFontStyle(const std::string& style_key,
                          const std::string& font_name = "Calibri",
                          double font_size = 11.0,
                          bool bold = false,
                          bool italic = false,
                          uint32_t color = 0x000000);
    
    /**
     * @brief 创建并缓存填充样式
     */
    size_t createFillStyle(const std::string& style_key,
                          PatternType pattern = PatternType::Solid,
                          uint32_t bg_color = 0xFFFFFF,
                          uint32_t fg_color = 0x000000);
    
    /**
     * @brief 创建并缓存边框样式
     */
    size_t createBorderStyle(const std::string& style_key,
                            BorderStyle style = BorderStyle::Thin,
                            uint32_t color = 0x000000);
    
    /**
     * @brief 创建组合样式
     */
    size_t createCompositeStyle(const std::string& style_key,
                               const std::string& font_key = "",
                               const std::string& fill_key = "",
                               const std::string& border_key = "");
    
    /**
     * @brief 获取格式仓储（用于样式操作）
     */
    FormatRepository* getFormatRepository() { return format_repo_.get(); }
    
    /**
     * @brief 获取样式模板（用于高级操作）
     */
    StyleTemplate* getStyleTemplate() { return template_.get(); }
    
    /**
     * @brief 清除所有缓存
     */
    void clearCache();
    
    /**
     * @brief 获取统计信息
     */
    struct Statistics {
        size_t imported_styles_count;
        size_t predefined_styles_count;
        size_t cached_styles_count;
        size_t total_format_pool_size;
    };
    
    Statistics getStatistics() const;
};

}} // namespace fastexcel::core