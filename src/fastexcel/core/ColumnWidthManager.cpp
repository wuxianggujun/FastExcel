#include "ColumnWidthManager.hpp"
#include "FormatRepository.hpp"
#include "StyleBuilder.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <algorithm>
#include <sstream>

namespace fastexcel {
namespace core {

ColumnWidthManager::ColumnWidthManager(FormatRepository* format_repo)
    : format_repo_(format_repo) {
    // 预热常用字体的MDW缓存
    mdw_cache_["Calibri"] = 7;
    mdw_cache_["Calibri_11"] = 7;
    mdw_cache_["Arial"] = 7;
    mdw_cache_["Arial_11"] = 7;
    mdw_cache_["微软雅黑"] = 8;     // 常见环境下≈8px
    mdw_cache_["微软雅黑_11"] = 8;
    mdw_cache_["宋体"] = 8;
    mdw_cache_["宋体_11"] = 8;
}

std::pair<double, int> ColumnWidthManager::setColumnWidth(int col, const ColumnWidthConfig& config) {
    // 1. 确定最终使用的字体
    std::string final_font = config.font_name;
    double final_size = config.font_size;
    
    if (config.strategy == WidthStrategy::ADAPTIVE && final_font.empty()) {
        // 自适应策略：根据目标宽度和内容推测最佳字体
        final_font = (config.target_width >= 5.0) ? "微软雅黑" : "Calibri";
        final_size = 11.0;
    }
    
    // 2. 计算最优列宽
    double optimal_width = calculateOptimalWidth(config.target_width, final_font, final_size);
    
    // 3. 获取或创建格式
    int format_id = -1;
    if (!final_font.empty() && final_font != "Calibri") {
        format_id = getOrCreateFontFormat(final_font, final_size);
    }
    
    FASTEXCEL_LOG_DEBUG("设置列{}: 目标={}, 优化={}, 字体={} {}pt, 格式ID={}", 
                       col, config.target_width, optimal_width, final_font, final_size, format_id);
    
    return {optimal_width, format_id};
}

std::unordered_map<int, std::pair<double, int>> ColumnWidthManager::setColumnWidths(
    const std::unordered_map<int, ColumnWidthConfig>& configs) {
    
    std::unordered_map<int, std::pair<double, int>> results;
    
    // 批量处理：按字体分组以提高效率
    std::unordered_map<std::string, std::vector<std::pair<int, ColumnWidthConfig>>> font_groups;
    
    for (const auto& [col, config] : configs) {
        std::string font_key = makeFontKey(config.font_name, config.font_size);
        font_groups[font_key].emplace_back(col, config);
    }
    
    // 逐组处理
    for (const auto& [font_key, group] : font_groups) {
        for (const auto& [col, config] : group) {
            results[col] = setColumnWidth(col, config);
        }
    }
    
    return results;
}

std::pair<double, int> ColumnWidthManager::setSmartColumnWidth(int col, double target_width,
                                                              const std::vector<std::string>& cell_contents) {
    // 智能分析单元格内容，选择最佳字体
    std::string optimal_font = selectOptimalFont(cell_contents);
    
    ColumnWidthConfig smart_config(target_width, optimal_font, 11.0, WidthStrategy::CONTENT_AWARE);
    return setColumnWidth(col, smart_config);
}

double ColumnWidthManager::calculateOptimalWidth(double target_width, const std::string& font_name, double font_size) const {
    // 根据指定字体选择合适的 MDW 值，而非固定使用 workbook_mdw_
    int effective_mdw = workbook_mdw_; // 默认使用工作簿MDW
    
    // 如果指定了特定字体，计算该字体的MDW
    if (!font_name.empty()) {
        effective_mdw = getMDW(font_name, font_size > 0 ? font_size : 11.0);
        FASTEXCEL_LOG_DEBUG("使用指定字体 {} {}pt 的 MDW: {}", font_name, font_size, effective_mdw);
    }
    
    auto& calculator = getCalculator(effective_mdw);
    double result = calculator.quantize(target_width);
    
    FASTEXCEL_LOG_DEBUG("列宽计算: 目标={} 字体={} MDW={} 结果={}", 
                       target_width, font_name.empty() ? "默认" : font_name, effective_mdw, result);
    
    return result;
}

int ColumnWidthManager::getOrCreateFontFormat(const std::string& font_name, double font_size) {
    if (!format_repo_) return -1;
    
    std::string font_key = makeFontKey(font_name, font_size);
    
    // 检查缓存
    auto it = format_cache_.find(font_key);
    if (it != format_cache_.end()) {
        return it->second;
    }
    
    // 创建新格式
    FormatDescriptor font_format = StyleBuilder()
        .fontName(font_name)
        .fontSize(font_size)
        .build();
    
    int format_id = format_repo_->addFormat(font_format);
    
    // 更新缓存
    format_cache_[font_key] = format_id;
    
    FASTEXCEL_LOG_DEBUG("创建字体格式: {} {}pt -> ID={}", font_name, font_size, format_id);
    
    return format_id;
}

void ColumnWidthManager::clearCache() {
    mdw_cache_.clear();
    format_cache_.clear();
    calculator_cache_.clear();
}

ColumnWidthManager::CacheStats ColumnWidthManager::getCacheStats() const {
    return {
        mdw_cache_.size(),
        format_cache_.size(),
        calculator_cache_.size()
    };
}

// 私有方法实现

int ColumnWidthManager::getMDW(const std::string& font_name, double font_size) const {
    std::string font_key = makeFontKey(font_name, font_size);
    
    auto it = mdw_cache_.find(font_key);
    if (it != mdw_cache_.end()) {
        return it->second;
    }
    
    // 计算并缓存
    int mdw = utils::ColumnWidthCalculator::estimateMDW(font_name, font_size);
    mdw_cache_[font_key] = mdw;
    
    return mdw;
}

utils::ColumnWidthCalculator& ColumnWidthManager::getCalculator(int mdw) const {
    auto it = calculator_cache_.find(mdw);
    if (it != calculator_cache_.end()) {
        return *it->second;
    }
    
    // 创建并缓存
    auto calculator = std::make_unique<utils::ColumnWidthCalculator>(mdw);
    auto& ref = *calculator;
    calculator_cache_[mdw] = std::move(calculator);
    
    return ref;
}

std::string ColumnWidthManager::makeFontKey(const std::string& font_name, double font_size) const {
    std::ostringstream oss;
    oss << font_name << "_" << font_size;
    return oss.str();
}

std::string ColumnWidthManager::selectOptimalFont(const std::vector<std::string>& contents) const {
    if (contents.empty()) return "Calibri";
    
    // 分析内容：统计中英文比例
    size_t chinese_count = 0;
    size_t total_chars = 0;
    
    for (const auto& content : contents) {
        for (unsigned char c : content) {
            ++total_chars;
            if (c >= 0x80) { // 简单的UTF-8中文检测
                ++chinese_count;
            }
        }
    }
    
    if (total_chars == 0) return "Calibri";
    
    double chinese_ratio = double(chinese_count) / double(total_chars);
    
    // 根据中文比例选择字体
    if (chinese_ratio > 0.3) {
        return "微软雅黑";  // 中文内容较多
    } else if (chinese_ratio > 0.1) {
        return "微软雅黑";  // 中英混合，倾向中文字体
    } else {
        return "Calibri";   // 主要是英文
    }
}

bool ColumnWidthManager::containsChinese(const std::string& text) const {
    for (unsigned char c : text) {
        if (c >= 0x80) { // 简单的UTF-8中文检测
            return true;
        }
    }
    return false;
}

}} // namespace fastexcel::core
