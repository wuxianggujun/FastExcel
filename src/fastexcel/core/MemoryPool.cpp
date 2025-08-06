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
        auto result = addNewBlock(block_size_);
        if (result.hasError()) {
            // 如果无法分配所有初始块，至少确保有一个块
            if (blocks_.empty()) {
                // 如果连一个块都分配不了，这是致命错误
                // 但我们需要优雅地处理这种情况
                return;
            }
            break;
        }
    }
}

MemoryPool::~MemoryPool() {
    auto result = clear();
    // 析构函数中忽略错误
}

Result<void*> MemoryPool::allocate(size_t size) {
    if (size == 0) {
        return makeError(ErrorCode::InvalidArgument, "Cannot allocate zero bytes");
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
        size_t new_block_size = std::max(size, block_size_);
        auto result = addNewBlock(new_block_size);
        if (result.hasError()) {
            // 内存分配失败，回滚统计信息
            stats_.allocation_count--;
            stats_.total_allocated -= size;
            stats_.current_usage -= size;
            return makeError(ErrorCode::OutOfMemory, "Failed to allocate new memory block");
        }
        block = blocks_.back().get();
    }
    
    block->in_use = true;
    return makeExpected(block->data);
}

VoidResult MemoryPool::deallocate(void* ptr) {
    if (!ptr) {
        return makeError(ErrorCode::InvalidArgument, "Cannot deallocate null pointer");
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
            return VoidResult();
        }
    }
    
    // 如果没找到对应的块，可能是外部分配的内存
    // 这种情况下直接释放
    std::free(ptr);
    return VoidResult();
}

VoidResult MemoryPool::clear() {
    std::lock_guard<std::mutex> lock(mutex_);
    blocks_.clear();
    stats_ = {};
    return VoidResult();
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

VoidResult MemoryPool::addNewBlock(size_t size) {
    auto result = Block::create(size);
    if (result.hasError()) {
        return makeError(result.error().code, "Failed to create new memory block: " + result.error().message);
    }
    blocks_.emplace_back(std::move(result.value()));
    return VoidResult();
}

// MemoryManager 实现
MemoryManager::MemoryManager()
    : default_pool_(std::make_unique<MemoryPool>(1024, 16)) {
}

MemoryManager::~MemoryManager() {
    auto result = cleanup();
    // 析构函数中忽略错误
}

MemoryManager& MemoryManager::getInstance() {
    static MemoryManager instance;
    return instance;
}

MemoryPool& MemoryManager::getDefaultPool() {
    return *default_pool_;
}

Result<std::reference_wrapper<MemoryPool>> MemoryManager::getPool(size_t block_size) {
    std::lock_guard<std::mutex> lock(mutex_);
    
    // 查找现有的池
    for (auto& [size, pool] : pools_) {
        if (size == block_size) {
            return makeExpected(std::ref(*pool));
        }
    }
    
    // 创建新的池
    auto new_pool = std::make_unique<MemoryPool>(block_size, 8);
    if (!new_pool) {
        return makeError(ErrorCode::OutOfMemory, "Failed to create new memory pool");
    }
    
    MemoryPool* pool_ptr = new_pool.get();
    pools_.emplace_back(block_size, std::move(new_pool));
    
    return makeExpected(std::ref(*pool_ptr));
}

VoidResult MemoryManager::cleanup() {
    std::lock_guard<std::mutex> lock(mutex_);
    pools_.clear();
    if (default_pool_) {
        auto result = default_pool_->clear();
        if (result.hasError()) {
            return result;
        }
    }
    return VoidResult();
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