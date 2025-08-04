/**
 * @file MemoryPool.cpp
 * @brief 内存池实现
 */

#include "MemoryPool.hpp"
#include <algorithm>
#include <stdexcept>
#include <cstdlib>

namespace fastexcel {
namespace core {

// MemoryPool 实现
MemoryPool::MemoryPool(size_t block_size, size_t initial_blocks)
    : block_size_(std::max(block_size, sizeof(void*)))
    , stats_{} {
    
    // 预分配初始内存块
    blocks_.reserve(initial_blocks * 2); // 预留更多空间避免频繁重分配
    
    for (size_t i = 0; i < initial_blocks; ++i) {
        try {
            addNewBlock(block_size_);
        } catch (const std::bad_alloc&) {
            // 如果无法分配所有初始块，至少确保有一个块
            if (blocks_.empty()) {
                throw;
            }
            break;
        }
    }
}

MemoryPool::~MemoryPool() {
    clear();
}

void* MemoryPool::allocate(size_t size) {
    if (size == 0) {
        return nullptr;
    }
    
    std::lock_guard<std::mutex> lock(mutex_);
    
    // 更新统计信息
    stats_.allocation_count++;
    stats_.total_allocated += size;
    stats_.current_usage += size;
    stats_.peak_usage = std::max(stats_.peak_usage, stats_.current_usage);
    
    // 查找可用的块
    Block* block = findAvailableBlock(size);
    if (!block) {
        // 没有可用块，创建新块
        try {
            size_t new_block_size = std::max(size, block_size_);
            addNewBlock(new_block_size);
            block = blocks_.back().get();
        } catch (const std::bad_alloc&) {
            // 内存分配失败，回滚统计信息
            stats_.allocation_count--;
            stats_.total_allocated -= size;
            stats_.current_usage -= size;
            return nullptr;
        }
    }
    
    block->in_use = true;
    return block->data;
}

void MemoryPool::deallocate(void* ptr) {
    if (!ptr) {
        return;
    }
    
    std::lock_guard<std::mutex> lock(mutex_);
    
    // 查找对应的块
    for (auto& block : blocks_) {
        if (block->data == ptr) {
            if (block->in_use) {
                block->in_use = false;
                stats_.deallocation_count++;
                stats_.total_deallocated += block->size;
                stats_.current_usage -= block->size;
            }
            return;
        }
    }
    
    // 如果没找到对应的块，可能是外部分配的内存
    // 这种情况下直接释放
    std::free(ptr);
}

void MemoryPool::clear() {
    std::lock_guard<std::mutex> lock(mutex_);
    blocks_.clear();
    stats_ = {};
}

MemoryPool::Statistics MemoryPool::getStatistics() const {
    std::lock_guard<std::mutex> lock(mutex_);
    return stats_;
}

void MemoryPool::resetStatistics() {
    std::lock_guard<std::mutex> lock(mutex_);
    stats_ = {};
}

MemoryPool::Block* MemoryPool::findAvailableBlock(size_t size) {
    for (auto& block : blocks_) {
        if (!block->in_use && block->size >= size) {
            return block.get();
        }
    }
    return nullptr;
}

void MemoryPool::addNewBlock(size_t size) {
    blocks_.emplace_back(std::make_unique<Block>(size));
}

// MemoryManager 实现
MemoryManager::MemoryManager()
    : default_pool_(std::make_unique<MemoryPool>(1024, 16)) {
}

MemoryManager::~MemoryManager() {
    cleanup();
}

MemoryManager& MemoryManager::getInstance() {
    static MemoryManager instance;
    return instance;
}

MemoryPool& MemoryManager::getDefaultPool() {
    return *default_pool_;
}

MemoryPool& MemoryManager::getPool(size_t block_size) {
    std::lock_guard<std::mutex> lock(mutex_);
    
    // 查找现有的池
    for (auto& [size, pool] : pools_) {
        if (size == block_size) {
            return *pool;
        }
    }
    
    // 创建新的池
    auto new_pool = std::make_unique<MemoryPool>(block_size, 8);
    MemoryPool* pool_ptr = new_pool.get();
    pools_.emplace_back(block_size, std::move(new_pool));
    
    return *pool_ptr;
}

void MemoryManager::cleanup() {
    std::lock_guard<std::mutex> lock(mutex_);
    pools_.clear();
    if (default_pool_) {
        default_pool_->clear();
    }
}

MemoryManager::GlobalStatistics MemoryManager::getGlobalStatistics() const {
    std::lock_guard<std::mutex> lock(mutex_);
    
    GlobalStatistics global_stats{};
    global_stats.total_pools = pools_.size() + (default_pool_ ? 1 : 0);
    
    if (default_pool_) {
        global_stats.default_pool_stats = default_pool_->getStatistics();
        global_stats.total_memory_allocated += global_stats.default_pool_stats.total_allocated;
        global_stats.total_memory_in_use += global_stats.default_pool_stats.current_usage;
    }
    
    for (const auto& [size, pool] : pools_) {
        auto stats = pool->getStatistics();
        global_stats.total_memory_allocated += stats.total_allocated;
        global_stats.total_memory_in_use += stats.current_usage;
    }
    
    return global_stats;
}

} // namespace core
} // namespace fastexcel