#pragma once

#include "fastexcel/utils/ColumnWidthCalculator.hpp"
#include "fastexcel/core/FormatDescriptor.hpp"
#include <string>
#include <unordered_map>
#include <memory>

namespace fastexcel {
namespace core {

// 前向声明
class FormatRepository;

/**
 * @brief 列宽管理器 - 统一管理列宽与字体的协调
 * 
 * 核心设计理念：
 * 1. 字体与列宽的一体化管理
 * 2. 智能缓存，避免重复计算
 * 3. 高性能批量操作
 * 4. 完全向后兼容
 */
class ColumnWidthManager {
public:
    /**
     * @brief 列宽配置策略
     */
    enum class WidthStrategy {
        EXACT,              // 精确匹配：根据内容字体计算最优MDW
        ADAPTIVE,           // 自适应：根据单元格内容自动选择字体
        CONTENT_AWARE,      // 内容感知：分析列内容分布，选择最佳字体
        LEGACY              // 兼容模式：使用传统的单一MDW计算
    };

    /**
     * @brief 列宽设置信息
     */
    struct ColumnWidthConfig {
        double target_width;           // 目标列宽
        std::string font_name;         // 字体名称
        double font_size;              // 字体大小
        WidthStrategy strategy;        // 策略
        int format_id = -1;           // 对应的格式ID
        bool is_optimized = false;    // 是否已优化
        
        ColumnWidthConfig(double width, const std::string& font = "", 
                         double size = 11.0, WidthStrategy strat = WidthStrategy::ADAPTIVE)
            : target_width(width), font_name(font), font_size(size), strategy(strat) {}
    };

private:
    FormatRepository* format_repo_;
    int workbook_mdw_ = 7; // 与Excel Normal(latin)字体对齐的MDW（默认7=Calibri 11）
    
    // 字体MDW缓存
    mutable std::unordered_map<std::string, int> mdw_cache_;
    
    // 格式ID缓存 (font_name + size -> format_id)
    mutable std::unordered_map<std::string, int> format_cache_;
    
    // 列宽计算器缓存
    mutable std::unordered_map<int, std::unique_ptr<utils::ColumnWidthCalculator>> calculator_cache_;

public:
    explicit ColumnWidthManager(FormatRepository* format_repo);
    
    // 设置/获取 工作簿Normal字体的MDW（用于列宽量化，确保与Excel一致）
    void setWorkbookNormalMDW(int mdw) { workbook_mdw_ = mdw; }
    int getWorkbookNormalMDW() const { return workbook_mdw_; }
    
    /**
     * @brief 设置列宽 - 新的统一接口
     * @param col 列号
     * @param config 列宽配置
     * @return 实际设置的列宽值和格式ID
     */
    std::pair<double, int> setColumnWidth(int col, const ColumnWidthConfig& config);
    
    /**
     * @brief 批量设置列宽 - 高性能批量操作
     * @param configs 列号到配置的映射
     * @return 实际设置结果
     */
    std::unordered_map<int, std::pair<double, int>> setColumnWidths(
        const std::unordered_map<int, ColumnWidthConfig>& configs);
    
    /**
     * @brief 智能列宽设置 - 根据单元格内容自动选择最佳字体和宽度
     * @param col 列号
     * @param target_width 目标宽度
     * @param cell_contents 该列的单元格内容样本
     * @return 优化后的列宽和格式ID
     */
    std::pair<double, int> setSmartColumnWidth(int col, double target_width,
                                              const std::vector<std::string>& cell_contents);
    
    /**
     * @brief 预计算列宽 - 不实际设置，只返回计算结果
     * @param target_width 目标宽度
     * @param font_name 字体名称
     * @param font_size 字体大小
     * @return 实际列宽值
     */
    double calculateOptimalWidth(double target_width, const std::string& font_name, double font_size) const;
    
    /**
     * @brief 获取或创建字体格式
     * @param font_name 字体名称
     * @param font_size 字体大小
     * @return 格式ID
     */
    int getOrCreateFontFormat(const std::string& font_name, double font_size);
    
    /**
     * @brief 清理缓存
     */
    void clearCache();
    
    /**
     * @brief 获取缓存统计信息
     */
    struct CacheStats {
        size_t mdw_cache_size;
        size_t format_cache_size;
        size_t calculator_cache_size;
    };
    CacheStats getCacheStats() const;

private:
    // 内部辅助方法
    int getMDW(const std::string& font_name, double font_size) const;
    utils::ColumnWidthCalculator& getCalculator(int mdw) const;
    std::string makeFontKey(const std::string& font_name, double font_size) const;
    
    // 智能字体选择
    std::string selectOptimalFont(const std::vector<std::string>& contents) const;
    bool containsChinese(const std::string& text) const;
};

}} // namespace fastexcel::core
