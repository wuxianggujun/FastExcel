#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/core/SharedFormula.hpp"
#include "fastexcel/utils/CommonUtils.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <regex>
#include <sstream>
#include <algorithm>

namespace fastexcel {
namespace core {

// SharedFormula 实现

SharedFormula::SharedFormula(int shared_index, const std::string& base_formula, const std::string& ref_range)
    : shared_index_(shared_index), base_formula_(base_formula), ref_range_(ref_range)
    , ref_first_row_(-1), ref_first_col_(-1), ref_last_row_(-1), ref_last_col_(-1) {
    parseReferenceRange(ref_range);
}

bool SharedFormula::isInRange(int row, int col) const {
    return (row >= ref_first_row_ && row <= ref_last_row_ &&
            col >= ref_first_col_ && col <= ref_last_col_);
}

std::string SharedFormula::expandFormula(int row, int col) const {
    if (!isInRange(row, col)) {
        FASTEXCEL_LOG_WARN("请求展开公式的位置({},{})不在共享公式范围内", row, col);
        return "";
    }
    
    return adjustFormula(base_formula_, ref_first_row_, ref_first_col_, row, col);
}

void SharedFormula::addAffectedCell(int row, int col) {
    if (isInRange(row, col)) {
        affected_cells_.emplace_back(row, col);
    }
}

SharedFormula::Statistics SharedFormula::getStatistics() const {
    Statistics stats;
    stats.affected_cells_count = affected_cells_.size();
    
    // 估算节省的内存：每个公式平均40字节，减去共享公式的存储开销
    size_t formula_size = base_formula_.size() + 40; // 基础公式 + 管理开销
    size_t individual_formulas_size = stats.affected_cells_count * (base_formula_.size() + 10);
    
    if (individual_formulas_size > formula_size) {
        stats.memory_saved = individual_formulas_size - formula_size;
        stats.compression_ratio = static_cast<double>(individual_formulas_size) / formula_size;
    }
    
    return stats;
}

void SharedFormula::parseReferenceRange(const std::string& range_str) {
    // 解析 "A1:C10" 格式的范围字符串
    size_t colon_pos = range_str.find(':');
    if (colon_pos == std::string::npos) {
        // 单个单元格引用
        auto [row, col] = utils::CommonUtils::parseReference(range_str);
        ref_first_row_ = ref_last_row_ = row;
        ref_first_col_ = ref_last_col_ = col;
    } else {
        // 范围引用
        std::string first_ref = range_str.substr(0, colon_pos);
        std::string last_ref = range_str.substr(colon_pos + 1);
        
        auto [first_row, first_col] = utils::CommonUtils::parseReference(first_ref);
        auto [last_row, last_col] = utils::CommonUtils::parseReference(last_ref);
        
        ref_first_row_ = first_row;
        ref_first_col_ = first_col;
        ref_last_row_ = last_row;
        ref_last_col_ = last_col;
    }
}

std::string SharedFormula::adjustFormula(const std::string& formula, int base_row, int base_col,
                                        int target_row, int target_col) const {
    // 使用正则表达式查找和替换单元格引用
    std::regex cell_ref_pattern(R"([A-Z]+[0-9]+)");
    std::string result = formula;
    
    int row_offset = target_row - base_row;
    int col_offset = target_col - base_col;
    
    if (row_offset == 0 && col_offset == 0) {
        return result; // 无需调整
    }
    
    // 查找所有单元格引用并调整
    std::sregex_iterator iter(formula.begin(), formula.end(), cell_ref_pattern);
    std::sregex_iterator end;
    
    // 从后往前替换，避免位置偏移问题
    std::vector<std::pair<size_t, std::pair<size_t, std::string>>> replacements;
    
    for (auto it = iter; it != end; ++it) {
        std::smatch match = *it;
        std::string cell_ref = match.str();
        
        try {
            auto [row, col] = utils::CommonUtils::parseReference(cell_ref);
            
            // 应用偏移
            int new_row = row + row_offset;
            int new_col = col + col_offset;
            
            // 确保新位置有效
            if (new_row >= 0 && new_col >= 0) {
                std::string new_ref = utils::CommonUtils::cellReference(new_row, new_col);
                replacements.emplace_back(match.position(), 
                                        std::make_pair(match.length(), new_ref));
            }
        } catch (const std::exception& e) {
            FASTEXCEL_LOG_WARN("解析单元格引用失败: {} ({})", cell_ref, e.what());
        }
    }
    
    // 从后往前执行替换
    std::sort(replacements.begin(), replacements.end(), std::greater<>());
    for (const auto& [pos, repl] : replacements) {
        result.replace(pos, repl.first, repl.second);
    }
    
    return result;
}

// SharedFormulaManager 实现

bool SharedFormulaManager::registerSharedFormula(const SharedFormula& shared_formula) {
    int index = shared_formula.getSharedIndex();
    
    // 检查索引是否已被使用
    if (shared_formulas_.find(index) != shared_formulas_.end()) {
        FASTEXCEL_LOG_WARN("共享公式索引 {} 已存在，将覆盖", index);
    }
    
    // 注册共享公式
    shared_formulas_[index] = shared_formula;
    
    // 更新单元格映射
    for (int row = shared_formula.getRefFirstRow(); row <= shared_formula.getRefLastRow(); ++row) {
        for (int col = shared_formula.getRefFirstCol(); col <= shared_formula.getRefLastCol(); ++col) {
            cell_to_shared_index_[std::make_pair(row, col)] = index;
        }
    }
    
    // 更新下一个可用索引
    if (index >= next_shared_index_) {
        next_shared_index_ = index + 1;
    }
    
    FASTEXCEL_LOG_DEBUG("注册共享公式成功: 索引={}, 范围={}", index, shared_formula.getRefRange());
    return true;
}

int SharedFormulaManager::registerSharedFormula(const std::string& base_formula, const std::string& ref_range) {
    int index = next_shared_index_++;
    SharedFormula shared_formula(index, base_formula, ref_range);
    
    if (registerSharedFormula(shared_formula)) {
        return index;
    }
    
    return -1; // 注册失败
}

int SharedFormulaManager::getSharedIndex(int row, int col) const {
    auto it = cell_to_shared_index_.find(std::make_pair(row, col));
    return (it != cell_to_shared_index_.end()) ? it->second : -1;
}

std::string SharedFormulaManager::getExpandedFormula(int row, int col) const {
    int shared_index = getSharedIndex(row, col);
    if (shared_index < 0) {
        return ""; // 不属于任何共享公式
    }
    
    auto it = shared_formulas_.find(shared_index);
    if (it == shared_formulas_.end()) {
        FASTEXCEL_LOG_ERROR("找不到共享公式索引: {}", shared_index);
        return "";
    }
    
    return it->second.expandFormula(row, col);
}

const SharedFormula* SharedFormulaManager::getSharedFormula(int shared_index) const {
    auto it = shared_formulas_.find(shared_index);
    return (it != shared_formulas_.end()) ? &it->second : nullptr;
}

bool SharedFormulaManager::isMainCell(int row, int col) const {
    int shared_index = getSharedIndex(row, col);
    if (shared_index < 0) {
        return false;
    }
    
    const SharedFormula* formula = getSharedFormula(shared_index);
    if (!formula) {
        return false;
    }
    
    // 主单元格通常是范围的第一个单元格（左上角）
    return (row == formula->getRefFirstRow() && col == formula->getRefFirstCol());
}

std::vector<SharedFormulaManager::FormulaPattern> SharedFormulaManager::detectSharedFormulaPatterns(
    const std::map<std::pair<int, int>, std::string>& formulas) const {
    
    std::vector<FormulaPattern> patterns;
    std::map<std::string, std::vector<std::pair<int, int>>> pattern_groups;
    
    // 按公式模式分组
    for (const auto& [pos, formula] : formulas) {
        std::string pattern = generateFormulaPattern(formula, pos.first, pos.second);
        pattern_groups[pattern].push_back(pos);
    }
    
    // 识别有效的共享公式模式（至少3个相似公式）
    for (const auto& [pattern, positions] : pattern_groups) {
        if (positions.size() >= 3) {
            FormulaPattern fp;
            fp.pattern_template = pattern;
            fp.matching_cells = positions;
            fp.estimated_savings = static_cast<int>(positions.size() * pattern.size() * 0.8); // 估算节省字节数
            patterns.push_back(fp);
        }
    }
    
    // 按节省空间排序
    std::sort(patterns.begin(), patterns.end(), 
              [](const FormulaPattern& a, const FormulaPattern& b) {
                  return a.estimated_savings > b.estimated_savings;
              });
    
    return patterns;
}

int SharedFormulaManager::optimizeFormulas(const std::map<std::pair<int, int>, std::string>& formulas,
                                          int min_count) {
    auto patterns = detectSharedFormulaPatterns(formulas);
    int optimized_count = 0;
    
    for (const auto& pattern : patterns) {
        if (static_cast<int>(pattern.matching_cells.size()) >= min_count) {
            // 确定范围
            int min_row = pattern.matching_cells[0].first;
            int max_row = pattern.matching_cells[0].first;
            int min_col = pattern.matching_cells[0].second;
            int max_col = pattern.matching_cells[0].second;
            
            for (const auto& [row, col] : pattern.matching_cells) {
                min_row = std::min(min_row, row);
                max_row = std::max(max_row, row);
                min_col = std::min(min_col, col);
                max_col = std::max(max_col, col);
            }
            
            // 创建共享公式
            std::string range = utils::CommonUtils::cellReference(min_row, min_col) + ":" +
                               utils::CommonUtils::cellReference(max_row, max_col);
            
            // 使用第一个匹配公式作为基础公式
            auto first_pos = pattern.matching_cells[0];
            auto formula_it = formulas.find(first_pos);
            if (formula_it != formulas.end()) {
                int shared_index = registerSharedFormula(formula_it->second, range);
                if (shared_index >= 0) {
                    // 为 SharedFormula 对象添加受影响的单元格
                    SharedFormula* shared_formula = &shared_formulas_[shared_index];
                    for (const auto& [row, col] : pattern.matching_cells) {
                        shared_formula->addAffectedCell(row, col);
                    }
                    
                    optimized_count += static_cast<int>(pattern.matching_cells.size());
                    FASTEXCEL_LOG_DEBUG("创建共享公式: 索引={}, 模式={}, 单元格数={}", 
                             shared_index, pattern.pattern_template, pattern.matching_cells.size());
                }
            }
        }
    }
    
    return optimized_count;
}

SharedFormulaManager::Statistics SharedFormulaManager::getStatistics() const {
    Statistics stats;
    stats.total_shared_formulas = shared_formulas_.size();
    
    for (const auto& [index, formula] : shared_formulas_) {
        auto formula_stats = formula.getStatistics();
        stats.total_affected_cells += formula_stats.affected_cells_count;
        stats.memory_saved += formula_stats.memory_saved;
    }
    
    if (stats.total_shared_formulas > 0) {
        // 计算平均压缩比
        double total_ratio = 0.0;
        for (const auto& [index, formula] : shared_formulas_) {
            auto formula_stats = formula.getStatistics();
            total_ratio += formula_stats.compression_ratio;
        }
        stats.average_compression_ratio = total_ratio / stats.total_shared_formulas;
    }
    
    return stats;
}

void SharedFormulaManager::clear() {
    shared_formulas_.clear();
    cell_to_shared_index_.clear();
    next_shared_index_ = 0;
    FASTEXCEL_LOG_DEBUG("清空所有共享公式数据");
}

std::vector<int> SharedFormulaManager::getAllSharedIndices() const {
    std::vector<int> indices;
    indices.reserve(shared_formulas_.size());
    
    for (const auto& [index, formula] : shared_formulas_) {
        indices.push_back(index);
    }
    
    std::sort(indices.begin(), indices.end());
    return indices;
}

void SharedFormulaManager::debugPrint() const {
    FASTEXCEL_LOG_DEBUG("=== 共享公式管理器状态 ===");
    FASTEXCEL_LOG_DEBUG("共享公式总数: {}", shared_formulas_.size());
    FASTEXCEL_LOG_DEBUG("下一个可用索引: {}", next_shared_index_);
    
    for (const auto& [index, formula] : shared_formulas_) {
        auto stats = formula.getStatistics();
        FASTEXCEL_LOG_DEBUG("索引 {}: 范围={}, 公式='{}', 影响单元格={}, 内存节省={}字节, 压缩比={:.2f}", 
                 index, formula.getRefRange(), formula.getBaseFormula(),
                 stats.affected_cells_count, stats.memory_saved, stats.compression_ratio);
    }
    FASTEXCEL_LOG_DEBUG("========================");
}

std::string SharedFormulaManager::generateFormulaPattern(const std::string& formula, int base_row, int base_col) const {
    // 将具体的单元格引用替换为模式标记
    std::regex cell_ref_pattern(R"([A-Z]+[0-9]+)");
    std::string pattern = formula;
    
    std::sregex_iterator iter(formula.begin(), formula.end(), cell_ref_pattern);
    std::sregex_iterator end;
    
    std::vector<std::pair<size_t, std::pair<size_t, std::string>>> replacements;
    
    for (auto it = iter; it != end; ++it) {
        std::smatch match = *it;
        std::string cell_ref = match.str();
        
        try {
            auto [row, col] = utils::CommonUtils::parseReference(cell_ref);
            
            // 计算相对偏移并生成模式标记
            int row_offset = row - base_row;
            int col_offset = col - base_col;
            
            std::string pattern_token = "{R" + std::to_string(row_offset) + "C" + std::to_string(col_offset) + "}";
            replacements.emplace_back(match.position(), 
                                    std::make_pair(match.length(), pattern_token));
        } catch (const std::exception& e) {
            FASTEXCEL_LOG_WARN("生成公式模式时解析引用失败: {} ({})", cell_ref, e.what());
        }
    }
    
    // 从后往前执行替换
    std::sort(replacements.begin(), replacements.end(), std::greater<>());
    for (const auto& [pos, repl] : replacements) {
        pattern.replace(pos, repl.first, repl.second);
    }
    
    return pattern;
}

bool SharedFormulaManager::isFormulaPatternMatch(const std::string& formula1, std::pair<int, int> pos1,
                                                const std::string& formula2, std::pair<int, int> pos2) const {
    std::string pattern1 = generateFormulaPattern(formula1, pos1.first, pos1.second);
    std::string pattern2 = generateFormulaPattern(formula2, pos2.first, pos2.second);
    
    return pattern1 == pattern2;
}

}} // namespace fastexcel::core
