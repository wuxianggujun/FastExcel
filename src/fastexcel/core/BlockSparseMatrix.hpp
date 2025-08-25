/**
 * @file BlockSparseMatrix.hpp
 * @brief 高性能分块稀疏矩阵，用于优化Worksheet单元格存储
 */

#pragma once

#include "fastexcel/core/Cell.hpp"
#include <unordered_map>
#include <array>
#include <bitset>
#include <memory>
#include <cstdint>

namespace fastexcel {
namespace core {

/**
 * @brief 高性能分块稀疏矩阵
 * 
 * 核心优化思想：
 * 1. 将工作表分割成64x64的固定大小块
 * 2. 只为有数据的块分配内存
 * 3. 使用位图快速判断单元格是否存在
 * 4. 块内使用连续内存存储，提高局部性
 * 5. O(1)的块定位，O(1)的块内访问
 */
class BlockSparseMatrix {
public:
    // 块大小设置 - 64x64是经过优化的平衡值
    // 太小：块数量多，管理开销大
    // 太大：内存浪费多，局部性差
    static constexpr int BLOCK_SIZE = 64;
    static constexpr int CELLS_PER_BLOCK = BLOCK_SIZE * BLOCK_SIZE;
    
private:
    /**
     * @brief 单元格块结构
     * 
     * 每个块包含64x64=4096个单元格
     * 使用位图快速标记哪些单元格有数据
     */
    struct CellBlock {
        // 连续存储的单元格数组 - 提高内存局部性
        std::array<Cell, CELLS_PER_BLOCK> cells;
        
        // 位图标记单元格是否被占用 - 4096位 = 512字节
        std::bitset<CELLS_PER_BLOCK> occupied;
        
        // 块的基础坐标
        int base_row;
        int base_col;
        
        // 统计信息
        int occupied_count = 0;  // 已占用单元格数量
        
        CellBlock(int row, int col) : base_row(row), base_col(col) {}
        
        /**
         * @brief 获取单元格引用
         * @param row 全局行号
         * @param col 全局列号
         * @return Cell引用
         */
        Cell& getCell(int row, int col) {
            int local_row = row - base_row;
            int local_col = col - base_col;
            int index = local_row * BLOCK_SIZE + local_col;
            
            // 如果首次访问，标记为占用
            if (!occupied.test(index)) {
                occupied.set(index);
                occupied_count++;
            }
            
            return cells[index];
        }
        
        /**
         * @brief 检查单元格是否存在
         * @param row 全局行号  
         * @param col 全局列号
         * @return 是否存在
         */
        bool hasCell(int row, int col) const {
            int local_row = row - base_row;
            int local_col = col - base_col;
            int index = local_row * BLOCK_SIZE + local_col;
            return occupied.test(index);
        }
        
        /**
         * @brief 移除单元格
         * @param row 全局行号
         * @param col 全局列号
         */
        void removeCell(int row, int col) {
            int local_row = row - base_row;
            int local_col = col - base_col;
            int index = local_row * BLOCK_SIZE + local_col;
            
            if (occupied.test(index)) {
                occupied.reset(index);
                occupied_count--;
                cells[index].clear();  // 清空单元格内容
            }
        }
        
        /**
         * @brief 获取块中所有非空单元格
         * @return vector<pair<pair<int,int>, Cell*>> 坐标和单元格指针对
         */
        std::vector<std::pair<std::pair<int, int>, Cell*>> getAllCells() {
            std::vector<std::pair<std::pair<int, int>, Cell*>> result;
            result.reserve(occupied_count);
            
            for (int i = 0; i < CELLS_PER_BLOCK; ++i) {
                if (occupied.test(i)) {
                    int local_row = i / BLOCK_SIZE;
                    int local_col = i % BLOCK_SIZE;
                    int global_row = base_row + local_row;
                    int global_col = base_col + local_col;
                    
                    result.emplace_back(std::make_pair(global_row, global_col), &cells[i]);
                }
            }
            
            return result;
        }
        
        /**
         * @brief 检查块是否为空
         */
        bool isEmpty() const {
            return occupied_count == 0;
        }
    };
    
    // 使用unordered_map存储块，key是块坐标的编码
    std::unordered_map<uint64_t, std::unique_ptr<CellBlock>> blocks_;
    
    /**
     * @brief 将行列坐标编码为块坐标key
     * @param row 行号
     * @param col 列号  
     * @return 64位编码的块key
     */
    uint64_t getBlockKey(int row, int col) const {
        int block_row = row / BLOCK_SIZE;
        int block_col = col / BLOCK_SIZE;
        
        // 将32位行坐标和32位列坐标合并为64位key
        return (static_cast<uint64_t>(block_row) << 32) | static_cast<uint64_t>(block_col);
    }
    
    /**
     * @brief 获取或创建指定位置的块
     * @param row 行号
     * @param col 列号
     * @return 块指针
     */
    CellBlock* getOrCreateBlock(int row, int col) {
        uint64_t key = getBlockKey(row, col);
        auto it = blocks_.find(key);
        
        if (it == blocks_.end()) {
            // 创建新块
            int block_row = (row / BLOCK_SIZE) * BLOCK_SIZE;
            int block_col = (col / BLOCK_SIZE) * BLOCK_SIZE;
            
            auto block = std::make_unique<CellBlock>(block_row, block_col);
            CellBlock* block_ptr = block.get();
            
            blocks_[key] = std::move(block);
            return block_ptr;
        }
        
        return it->second.get();
    }
    
    /**
     * @brief 获取指定位置的块（只读）
     * @param row 行号
     * @param col 列号
     * @return 块指针，如果不存在返回nullptr
     */
    CellBlock* getBlock(int row, int col) const {
        uint64_t key = getBlockKey(row, col);
        auto it = blocks_.find(key);
        return it != blocks_.end() ? it->second.get() : nullptr;
    }
    
public:
    /**
     * @brief 默认构造函数
     */
    BlockSparseMatrix() = default;
    
    /**
     * @brief 析构函数
     */
    ~BlockSparseMatrix() = default;
    
    /**
     * @brief 获取单元格引用
     * @param row 行号
     * @param col 列号
     * @return Cell引用
     */
    Cell& getCell(int row, int col) {
        CellBlock* block = getOrCreateBlock(row, col);
        return block->getCell(row, col);
    }
    
    /**
     * @brief 获取单元格常量引用
     * @param row 行号
     * @param col 列号
     * @return Cell常量引用，如果不存在返回空Cell
     */
    const Cell& getCell(int row, int col) const {
        CellBlock* block = getBlock(row, col);
        if (block && block->hasCell(row, col)) {
            return const_cast<CellBlock*>(block)->getCell(row, col);
        }
        
        // 返回静态空Cell
        static Cell empty_cell;
        return empty_cell;
    }
    
    /**
     * @brief 检查指定位置是否有单元格
     * @param row 行号
     * @param col 列号
     * @return 是否存在单元格
     */
    bool hasCell(int row, int col) const {
        CellBlock* block = getBlock(row, col);
        return block && block->hasCell(row, col);
    }
    
    /**
     * @brief 移除指定位置的单元格
     * @param row 行号
     * @param col 列号
     */
    void removeCell(int row, int col) {
        CellBlock* block = getBlock(row, col);
        if (block) {
            block->removeCell(row, col);
            
            // 如果块变空，删除块以节省内存
            if (block->isEmpty()) {
                uint64_t key = getBlockKey(row, col);
                blocks_.erase(key);
            }
        }
    }
    
    /**
     * @brief 获取所有非空单元格
     * @return vector<pair<pair<int,int>, Cell*>> 坐标和单元格指针对
     */
    std::vector<std::pair<std::pair<int, int>, Cell*>> getAllCells() {
        std::vector<std::pair<std::pair<int, int>, Cell*>> result;
        
        // 预估大小以减少重新分配
        size_t total_cells = 0;
        for (const auto& [key, block] : blocks_) {
            total_cells += block->occupied_count;
        }
        result.reserve(total_cells);
        
        // 收集所有块中的单元格
        for (auto& [key, block] : blocks_) {
            auto block_cells = block->getAllCells();
            result.insert(result.end(), 
                         std::make_move_iterator(block_cells.begin()),
                         std::make_move_iterator(block_cells.end()));
        }
        
        return result;
    }
    
    /**
     * @brief 获取所有非空单元格（常量版本）
     */
    std::vector<std::pair<std::pair<int, int>, const Cell*>> getAllCells() const {
        std::vector<std::pair<std::pair<int, int>, const Cell*>> result;
        
        size_t total_cells = 0;
        for (const auto& [key, block] : blocks_) {
            total_cells += block->occupied_count;
        }
        result.reserve(total_cells);
        
        for (const auto& [key, block] : blocks_) {
            auto block_cells = const_cast<CellBlock*>(block.get())->getAllCells();
            for (const auto& [pos, cell] : block_cells) {
                result.emplace_back(pos, cell);
            }
        }
        
        return result;
    }
    
    /**
     * @brief 清空所有单元格
     */
    void clear() {
        blocks_.clear();
    }
    
    /**
     * @brief 获取非空单元格总数
     * @return 单元格数量
     */
    size_t size() const {
        size_t total = 0;
        for (const auto& [key, block] : blocks_) {
            total += block->occupied_count;
        }
        return total;
    }
    
    /**
     * @brief 检查矩阵是否为空
     * @return 是否为空
     */
    bool empty() const {
        return blocks_.empty();
    }
    
    /**
     * @brief 获取内存使用统计
     */
    struct MemoryStats {
        size_t total_blocks;        // 总块数
        size_t total_cells;         // 总单元格数
        size_t occupied_cells;      // 已占用单元格数
        size_t memory_usage_bytes;  // 内存使用量（字节）
        double occupancy_rate;      // 占用率
    };
    
    MemoryStats getMemoryStats() const {
        MemoryStats stats = {};
        
        stats.total_blocks = blocks_.size();
        stats.total_cells = stats.total_blocks * CELLS_PER_BLOCK;
        
        for (const auto& [key, block] : blocks_) {
            stats.occupied_cells += block->occupied_count;
        }
        
        stats.memory_usage_bytes = stats.total_blocks * sizeof(CellBlock);
        stats.occupancy_rate = stats.total_cells > 0 ? 
            static_cast<double>(stats.occupied_cells) / stats.total_cells : 0.0;
            
        return stats;
    }
    
    // 禁用拷贝构造和赋值（可根据需要实现）
    BlockSparseMatrix(const BlockSparseMatrix&) = delete;
    BlockSparseMatrix& operator=(const BlockSparseMatrix&) = delete;
    
    // 支持移动语义
    BlockSparseMatrix(BlockSparseMatrix&&) = default;
    BlockSparseMatrix& operator=(BlockSparseMatrix&&) = default;
};

}} // namespace fastexcel::core