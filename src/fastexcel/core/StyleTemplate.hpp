#pragma once

#include "fastexcel/core/Format.hpp"
#include <memory>
#include <unordered_map>
#include <string>
#include <vector>

namespace fastexcel {
namespace core {

/**
 * @brief 样式模板类 - 提供预设样式和可复用的样式配置
 */
class StyleTemplate {
private:
    std::unordered_map<std::string, std::shared_ptr<Format>> predefined_styles_;
    std::unordered_map<int, std::shared_ptr<Format>> imported_styles_; // 从文件导入的样式
    
public:
    StyleTemplate();
    ~StyleTemplate() = default;

    // 预定义样式创建
    void createPredefinedStyles();
    
    /**
     * @brief 获取预定义样式
     * @param name 样式名称，如 "header", "data", "number", "currency" 等
     * @return 样式对象指针
     */
    std::shared_ptr<Format> getPredefinedStyle(const std::string& name) const;
    
    /**
     * @brief 添加自定义样式
     * @param name 样式名称
     * @param format 样式对象
     */
    void addCustomStyle(const std::string& name, std::shared_ptr<Format> format);
    
    /**
     * @brief 从文件导入样式
     * @param styles 样式映射（索引->格式）
     */
    void importStylesFromFile(const std::unordered_map<int, std::shared_ptr<Format>>& styles);
    
    /**
     * @brief 获取导入的样式
     * @param index 原始样式索引
     * @return 样式对象指针
     */
    std::shared_ptr<Format> getImportedStyle(int index) const;
    
    /**
     * @brief 获取所有导入的样式
     */
    const std::unordered_map<int, std::shared_ptr<Format>>& getImportedStyles() const {
        return imported_styles_;
    }
    
    /**
     * @brief 创建字体样式
     * @param font_name 字体名称
     * @param font_size 字体大小
     * @param bold 是否粗体
     * @param italic 是否斜体
     * @param color 字体颜色
     * @return 字体样式
     */
    std::shared_ptr<Format> createFontStyle(
        const std::string& font_name = "Calibri",
        double font_size = 11.0,
        bool bold = false,
        bool italic = false,
        uint32_t color = 0x000000
    );
    
    /**
     * @brief 创建填充样式
     * @param pattern 填充模式
     * @param bg_color 背景色
     * @param fg_color 前景色
     * @return 填充样式
     */
    std::shared_ptr<Format> createFillStyle(
        PatternType pattern = PatternType::Solid,
        uint32_t bg_color = 0xFFFFFF,
        uint32_t fg_color = 0x000000
    );
    
    /**
     * @brief 创建边框样式
     * @param style 边框样式
     * @param color 边框颜色
     * @return 边框样式
     */
    std::shared_ptr<Format> createBorderStyle(
        BorderStyle style = BorderStyle::Thin,
        uint32_t color = 0x000000
    );
    
    /**
     * @brief 创建组合样式
     * @param font_style 字体样式
     * @param fill_style 填充样式
     * @param border_style 边框样式
     * @return 组合样式
     */
    std::shared_ptr<Format> createCompositeStyle(
        std::shared_ptr<Format> font_style = nullptr,
        std::shared_ptr<Format> fill_style = nullptr,
        std::shared_ptr<Format> border_style = nullptr
    );
};

}} // namespace fastexcel::core