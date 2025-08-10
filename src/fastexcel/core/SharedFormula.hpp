#pragma once

#include <string>
#include <vector>
#include <map>
#include <utility>
#include <memory>

namespace fastexcel {
namespace core {

/**
 * @brief 共享公式类型
 */
enum class SharedFormulaType {
    NORMAL,     // 普通公式
    SHARED,     // 共享公式主公式
    REFERENCE   // 共享公式引用
};

/**
 * @brief 共享公式数据结构
 * 
 * Excel中的共享公式(Shared Formulas)是一种优化技术，用于存储大量相似的公式。
 * 例如：A1:A100都包含类似"=B1*C1", "=B2*C2"的公式，Excel会将其存储为一个共享公式。
 */
class SharedFormula {
private:
    int shared_index_;                                 // 共享公式索引 (si属性)
    std::string base_formula_;                        // 基础公式模板
    std::string ref_range_;                           // 应用范围 (如"A1:A100")
    int ref_first_row_, ref_first_col_;              // 范围起始位置
    int ref_last_row_, ref_last_col_;               // 范围结束位置
    std::vector<std::pair<int, int>> affected_cells_; // 受影响的单元格列表

public:
    /**
     * @brief 构造函数
     * @param shared_index 共享公式索引
     * @param base_formula 基础公式
     * @param ref_range 引用范围字符串 (如"A1:C10")
     */
    SharedFormula(int shared_index, const std::string& base_formula, const std::string& ref_range);
    
    /**
     * @brief 默认构造函数
     */
    SharedFormula() : shared_index_(-1), ref_first_row_(-1), ref_first_col_(-1), 
                     ref_last_row_(-1), ref_last_col_(-1) {}
    
    // ========== Getters ==========
    
    int getSharedIndex() const { return shared_index_; }
    const std::string& getBaseFormula() const { return base_formula_; }
    const std::string& getRefRange() const { return ref_range_; }
    
    int getRefFirstRow() const { return ref_first_row_; }
    int getRefFirstCol() const { return ref_first_col_; }
    int getRefLastRow() const { return ref_last_row_; }
    int getRefLastCol() const { return ref_last_col_; }
    
    const std::vector<std::pair<int, int>>& getAffectedCells() const { return affected_cells_; }
    
    // ========== 核心功能 ==========
    
    /**
     * @brief 检查指定单元格是否在共享公式范围内
     * @param row 行号
     * @param col 列号
     * @return 是否在范围内
     */
    bool isInRange(int row, int col) const;
    
    /**
     * @brief 为指定单元格展开具体的公式
     * @param row 目标单元格行号
     * @param col 目标单元格列号
     * @return 展开后的具体公式
     * 
     * 例如：基础公式"=B1*C1"，对于位置(2,0)会展开为"=B3*C3"
     */
    std::string expandFormula(int row, int col) const;
    
    /**
     * @brief 添加受影响的单元格
     * @param row 行号
     * @param col 列号
     */
    void addAffectedCell(int row, int col);
    
    /**
     * @brief 获取统计信息
     */
    struct Statistics {
        size_t affected_cells_count = 0;
        size_t memory_saved = 0;      // 节省的内存字节数（估算）
        double compression_ratio = 0.0; // 压缩比
    };
    Statistics getStatistics() const;
    
    // ========== 内部工具方法 ==========
    
private:
    /**
     * @brief 解析范围字符串 (如"A1:C10")
     * @param range_str 范围字符串
     */
    void parseReferenceRange(const std::string& range_str);
    
    /**
     * @brief 调整公式中的单元格引用
     * @param formula 原始公式
     * @param base_row 基础行
     * @param base_col 基础列
     * @param target_row 目标行
     * @param target_col 目标列
     * @return 调整后的公式
     */
    std::string adjustFormula(const std::string& formula, int base_row, int base_col, 
                             int target_row, int target_col) const;
};

/**
 * @brief 共享公式管理器
 * 
 * 负责管理整个工作表中的所有共享公式，提供注册、查询、优化等功能。
 */
class SharedFormulaManager {
private:
    std::map<int, SharedFormula> shared_formulas_;           // 索引 -> 共享公式
    std::map<std::pair<int, int>, int> cell_to_shared_index_; // 单元格位置 -> 共享索引
    int next_shared_index_;                                   // 下一个可用的共享索引
    
public:
    SharedFormulaManager() : next_shared_index_(0) {}
    
    // ========== 注册和管理 ==========
    
    /**
     * @brief 注册一个共享公式
     * @param shared_formula 共享公式对象
     * @return 是否注册成功
     */
    bool registerSharedFormula(const SharedFormula& shared_formula);
    
    /**
     * @brief 自动分配共享索引并注册
     * @param base_formula 基础公式
     * @param ref_range 引用范围
     * @return 分配的共享索引，失败返回-1
     */
    int registerSharedFormula(const std::string& base_formula, const std::string& ref_range);
    
    // ========== 查询功能 ==========
    
    /**
     * @brief 检查指定单元格是否属于某个共享公式
     * @param row 行号
     * @param col 列号
     * @return 共享索引，不属于任何共享公式返回-1
     */
    int getSharedIndex(int row, int col) const;
    
    /**
     * @brief 根据单元格位置获取展开的公式
     * @param row 行号
     * @param col 列号
     * @return 展开的公式，如果不是共享公式返回空字符串
     */
    std::string getExpandedFormula(int row, int col) const;
    
    /**
     * @brief 根据共享索引获取共享公式对象
     * @param shared_index 共享索引
     * @return 共享公式指针，不存在返回nullptr
     */
    const SharedFormula* getSharedFormula(int shared_index) const;
    
    /**
     * @brief 检查指定单元格是否是共享公式的主单元格
     * @param row 行号
     * @param col 列号
     * @return 是否为主单元格
     */
    bool isMainCell(int row, int col) const;
    
    // ========== 优化功能 ==========
    
    /**
     * @brief 公式模式检测结果
     */
    struct FormulaPattern {
        std::string pattern_template;                    // 公式模板
        std::vector<std::pair<int, int>> matching_cells; // 匹配的单元格
        int estimated_savings;                           // 预估节省的字节数
    };
    
    /**
     * @brief 检测工作表中可以优化为共享公式的公式模式
     * @param formulas 单元格位置 -> 公式的映射
     * @return 检测到的模式列表
     */
    std::vector<FormulaPattern> detectSharedFormulaPatterns(
        const std::map<std::pair<int, int>, std::string>& formulas) const;
    
    /**
     * @brief 自动优化：将相似公式转换为共享公式
     * @param formulas 输入的公式映射
     * @param min_count 最小重复次数（少于此数不优化）
     * @return 优化的公式数量
     */
    int optimizeFormulas(const std::map<std::pair<int, int>, std::string>& formulas, 
                        int min_count = 3);
    
    // ========== 统计和诊断 ==========
    
    /**
     * @brief 获取共享公式统计信息
     */
    struct Statistics {
        size_t total_shared_formulas = 0;    // 共享公式总数
        size_t total_affected_cells = 0;     // 受影响的单元格总数
        size_t memory_saved = 0;             // 节省的内存（估算）
        double average_compression_ratio = 0.0; // 平均压缩比
    };
    Statistics getStatistics() const;
    
    /**
     * @brief 清空所有共享公式
     */
    void clear();
    
    /**
     * @brief 获取所有共享公式的索引列表
     */
    std::vector<int> getAllSharedIndices() const;
    
    /**
     * @brief 调试输出所有共享公式信息
     */
    void debugPrint() const;

private:
    /**
     * @brief 生成公式模式模板
     * @param formula 具体公式
     * @param base_row 基础行
     * @param base_col 基础列
     * @return 公式模板
     */
    std::string generateFormulaPattern(const std::string& formula, int base_row, int base_col) const;
    
    /**
     * @brief 检查两个公式是否匹配相同的模式
     * @param formula1 公式1
     * @param pos1 公式1的位置
     * @param formula2 公式2
     * @param pos2 公式2的位置
     * @return 是否匹配
     */
    bool isFormulaPatternMatch(const std::string& formula1, std::pair<int, int> pos1,
                              const std::string& formula2, std::pair<int, int> pos2) const;
};

}} // namespace fastexcel::core