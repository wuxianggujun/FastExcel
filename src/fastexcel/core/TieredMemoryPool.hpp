/**
 * @file TieredMemoryPool.hpp
 * @brief 分级内存池，优化不同大小内存块的分配性能
 */

#pragma once

#include "MemoryPool.hpp"
#include <array>
#include <vector>
#include <memory>
#include <mutex>
#include <cstddef>

namespace fastexcel {
namespace core {

/**
 * @brief 分级内存池实现
 * 
 * 核心优化思想：
 * 1. 将内存分配请求按大小分类到不同的级别
 * 2. 每个级别维护自己的空闲块列表，实现O(1)分配
 * 3. 使用预定义的大小级别，减少内存碎片
 * 4. 大块分配和小块分配分别优化
 */
class TieredMemoryPool {
public:
    /**
     * @brief 内存大小级别定义
     * 
     * 基于常见的内存使用模式设计：
     * - 32B: 小对象（如指针、基本结构体）
     * - 64B: 中小对象（如字符串、小数组）
     * - 128B: 中等对象（如Cell对象）
     * - 256B: 大对象（如复杂结构体）
     * - 512B, 1KB, 2KB, 4KB: 大块内存
     */
    static constexpr size_t SIZE_CLASS_COUNT = 8;
    static constexpr std::array<size_t, SIZE_CLASS_COUNT> SIZE_CLASSES = {
        32, 64, 128, 256, 512, 1024, 2048, 4096
    };
    
private:
    /**
     * @brief 单个大小级别的内存池
     */
    class SizeClass {
    private:
        size_t block_size_;                           // 这个级别的块大小
        std::vector<void*> free_blocks_;             // 空闲块列表
        std::vector<std::unique_ptr<char[]>> chunks_; // 内存块列表
        size_t blocks_per_chunk_;                    // 每个chunk包含的块数
        
        // 统计信息
        size_t allocated_blocks_ = 0;
        size_t total_blocks_ = 0;
        
    public:
        /**
         * @brief 构造函数
         * @param size 块大小
         */
        explicit SizeClass(size_t size) : block_size_(size) {
            // 计算每个chunk的块数 - 目标是每个chunk大约4KB
            blocks_per_chunk_ = std::max(size_t(1), 4096 / size);
        }
        
        /**
         * @brief 分配一个块
         * @return 分配的内存指针，失败返回nullptr
         */
        void* allocate() {
            if (free_blocks_.empty()) {
                if (!addNewChunk()) {
                    return nullptr; // 内存不足
                }
            }
            
            void* ptr = free_blocks_.back();
            free_blocks_.pop_back();
            allocated_blocks_++;
            
            return ptr;
        }
        
        /**
         * @brief 释放一个块
         * @param ptr 要释放的内存指针
         * @return 是否成功释放
         */
        bool deallocate(void* ptr) {
            if (!ptr) return false;
            
            // 验证指针是否属于这个SizeClass
            if (!isValidPointer(ptr)) {
                return false;
            }
            
            free_blocks_.push_back(ptr);
            allocated_blocks_--;
            
            return true;
        }
        
        /**
         * @brief 获取统计信息
         */
        struct Stats {
            size_t block_size;
            size_t allocated_blocks;
            size_t free_blocks;
            size_t total_blocks;
            size_t chunks_count;
            size_t memory_usage;
        };
        
        Stats getStats() const {
            return {
                block_size_,
                allocated_blocks_,
                free_blocks_.size(),
                total_blocks_,
                chunks_.size(),
                chunks_.size() * blocks_per_chunk_ * block_size_
            };
        }
        
        /**
         * @brief 预热 - 预分配指定数量的块
         * @param count 要预分配的块数量
         */
        void warmup(size_t count) {
            size_t chunks_needed = (count + blocks_per_chunk_ - 1) / blocks_per_chunk_;
            for (size_t i = 0; i < chunks_needed; ++i) {
                addNewChunk();
            }
        }
        
    private:
        /**
         * @brief 添加新的内存块
         * @return 是否成功
         */
        bool addNewChunk() {
            try {
                size_t chunk_size = block_size_ * blocks_per_chunk_;
                auto chunk = std::make_unique<char[]>(chunk_size);
                char* ptr = chunk.get();
                
                // 将chunk分割成多个块并添加到空闲列表
                for (size_t i = 0; i < blocks_per_chunk_; ++i) {
                    free_blocks_.push_back(ptr + i * block_size_);
                }
                
                total_blocks_ += blocks_per_chunk_;
                chunks_.push_back(std::move(chunk));
                
                return true;
            } catch (const std::exception&) {
                return false;
            }
        }
        
        /**
         * @brief 验证指针是否属于这个SizeClass
         * @param ptr 要验证的指针
         * @return 是否有效
         */
        bool isValidPointer(void* ptr) const {
            char* char_ptr = static_cast<char*>(ptr);
            
            for (const auto& chunk : chunks_) {
                char* chunk_start = chunk.get();
                char* chunk_end = chunk_start + (block_size_ * blocks_per_chunk_);
                
                if (char_ptr >= chunk_start && char_ptr < chunk_end) {
                    // 检查指针是否对齐到块边界
                    size_t offset = char_ptr - chunk_start;
                    return (offset % block_size_) == 0;
                }
            }
            
            return false;
        }
    };
    
    // 所有大小级别的内存池
    std::array<SizeClass, SIZE_CLASS_COUNT> size_classes_;
    
    // 线程安全保护
    mutable std::mutex mutex_;
    
    // 大块内存的后备分配器（超过最大SIZE_CLASS的请求）
    std::vector<std::unique_ptr<char[]>> large_blocks_;
    
public:
    /**
     * @brief 构造函数
     */
    TieredMemoryPool() {
        // 初始化每个大小级别
        for (size_t i = 0; i < SIZE_CLASSES.size(); ++i) {
            size_classes_[i] = SizeClass(SIZE_CLASSES[i]);
        }
    }
    
    /**
     * @brief 分配内存
     * @param size 需要分配的字节数
     * @return 分配的内存指针，失败返回nullptr
     */
    void* allocate(size_t size) {
        if (size == 0) return nullptr;
        
        std::lock_guard<std::mutex> lock(mutex_);
        
        // 查找合适的大小级别
        size_t class_index = getSizeClassIndex(size);
        
        if (class_index < SIZE_CLASS_COUNT) {
            // 使用对应的SizeClass分配
            return size_classes_[class_index].allocate();
        } else {
            // 大块内存直接分配
            return allocateLargeBlock(size);
        }
    }
    
    /**
     * @brief 释放内存
     * @param ptr 要释放的内存指针
     * @param size 原始分配的大小（用于确定级别）
     */
    void deallocate(void* ptr, size_t size) {
        if (!ptr) return;
        
        std::lock_guard<std::mutex> lock(mutex_);
        
        size_t class_index = getSizeClassIndex(size);
        
        if (class_index < SIZE_CLASS_COUNT) {
            // 使用对应的SizeClass释放
            if (!size_classes_[class_index].deallocate(ptr)) {
                // 如果SizeClass拒绝释放，可能是大块内存
                deallocateLargeBlock(ptr);
            }
        } else {
            // 大块内存直接释放
            deallocateLargeBlock(ptr);
        }
    }
    
    /**
     * @brief 获取内存池统计信息
     */
    struct PoolStats {
        std::array<SizeClass::Stats, SIZE_CLASS_COUNT> size_class_stats;
        size_t large_blocks_count;
        size_t total_memory_usage;
    };
    
    PoolStats getStats() const {
        std::lock_guard<std::mutex> lock(mutex_);
        
        PoolStats stats = {};
        stats.large_blocks_count = large_blocks_.size();
        stats.total_memory_usage = 0;
        
        for (size_t i = 0; i < SIZE_CLASS_COUNT; ++i) {
            stats.size_class_stats[i] = size_classes_[i].getStats();
            stats.total_memory_usage += stats.size_class_stats[i].memory_usage;
        }
        
        return stats;
    }
    
    /**
     * @brief 预热内存池
     * @param class_index 大小级别索引
     * @param count 预分配的块数量
     */
    void warmup(size_t class_index, size_t count) {
        if (class_index >= SIZE_CLASS_COUNT) return;
        
        std::lock_guard<std::mutex> lock(mutex_);
        size_classes_[class_index].warmup(count);
    }
    
    /**
     * @brief 预热所有级别
     * @param blocks_per_class 每个级别预分配的块数
     */
    void warmupAll(size_t blocks_per_class = 32) {
        for (size_t i = 0; i < SIZE_CLASS_COUNT; ++i) {
            warmup(i, blocks_per_class);
        }
    }
    
private:
    /**
     * @brief 获取大小对应的级别索引
     * @param size 请求的大小
     * @return 级别索引，如果超过最大级别返回SIZE_CLASS_COUNT
     */
    size_t getSizeClassIndex(size_t size) const {
        for (size_t i = 0; i < SIZE_CLASSES.size(); ++i) {
            if (size <= SIZE_CLASSES[i]) {
                return i;
            }
        }
        return SIZE_CLASS_COUNT; // 超过最大级别
    }
    
    /**
     * @brief 分配大块内存
     * @param size 请求的大小
     * @return 分配的内存指针
     */
    void* allocateLargeBlock(size_t size) {
        try {
            auto block = std::make_unique<char[]>(size);
            void* ptr = block.get();
            large_blocks_.push_back(std::move(block));
            return ptr;
        } catch (const std::exception&) {
            return nullptr;
        }
    }
    
    /**
     * @brief 释放大块内存
     * @param ptr 要释放的内存指针
     */
    void deallocateLargeBlock(void* ptr) {
        auto it = std::find_if(large_blocks_.begin(), large_blocks_.end(),
                              [ptr](const std::unique_ptr<char[]>& block) {
                                  return block.get() == ptr;
                              });
        
        if (it != large_blocks_.end()) {
            large_blocks_.erase(it);
        }
    }
    
    // 禁用拷贝构造和赋值
    TieredMemoryPool(const TieredMemoryPool&) = delete;
    TieredMemoryPool& operator=(const TieredMemoryPool&) = delete;
};

// 静态成员定义
constexpr std::array<size_t, TieredMemoryPool::SIZE_CLASS_COUNT> TieredMemoryPool::SIZE_CLASSES;

}} // namespace fastexcel::core