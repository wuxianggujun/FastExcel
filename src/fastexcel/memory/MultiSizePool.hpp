/**
 * @file MultiSizePool.hpp
 * @brief 多大小内存池实现
 */

#pragma once

#include "AlignedAllocator.hpp"
#include "IMemoryPool.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <vector>
#include <mutex>
#include <memory>
#include <functional>
#include <algorithm>

namespace fastexcel {
namespace memory {

/**
 * @brief 多大小内存池
 * 
 * 支持多种对象大小的内存分配，针对不同大小的对象进行优化
 */
class MultiSizePool : public IMemoryPool {
private:
    struct Page {
        void* buffer = nullptr;
        size_t stride = 0;
        size_t blocks = 0;
        ~Page() { AlignedAllocator::deallocate(buffer); }
    };

    struct SizeClass {
        size_t size = 0;        // 逻辑块大小（对外）
        size_t alignment = 0;   // 对齐
        size_t stride = 0;      // 内部步长（>= size，满足alignment）
        std::vector<std::unique_ptr<Page>> pages;
        std::vector<void*> free_list; // 当前可用块地址
        mutable std::mutex mtx;

        SizeClass(size_t s, size_t align)
            : size(s), alignment(align) {
            stride = ((size + alignment - 1) / alignment) * alignment;
            if (stride < size) stride = size; // 溢出保护
        }

        // 非拷贝
        SizeClass(const SizeClass&) = delete;
        SizeClass& operator=(const SizeClass&) = delete;
        // 自定义移动：互斥量不移动，目标重新构造
        SizeClass(SizeClass&& other) noexcept
            : size(other.size)
            , alignment(other.alignment)
            , stride(other.stride)
            , pages(std::move(other.pages))
            , free_list(std::move(other.free_list)) {
        }
        SizeClass& operator=(SizeClass&& other) noexcept {
            if (this != &other) {
                size = other.size;
                alignment = other.alignment;
                stride = other.stride;
                pages = std::move(other.pages);
                free_list = std::move(other.free_list);
            }
            return *this;
        }
    };

public:
    /**
     * @brief 构造函数
     */
    MultiSizePool() {
        // 初始化常用大小类
        initializeSizeClasses();
        
        FASTEXCEL_LOG_DEBUG("Created MultiSizePool with {} size classes", size_classes_.size());
    }
    
    /**
     * @brief 析构函数
     */
    ~MultiSizePool() = default;
    
    // 禁用拷贝和移动
    MultiSizePool(const MultiSizePool&) = delete;
    MultiSizePool& operator=(const MultiSizePool&) = delete;
    MultiSizePool(MultiSizePool&&) = delete;
    MultiSizePool& operator=(MultiSizePool&&) = delete;
    
    /**
     * @brief 分配指定大小的内存
     * 实现 IMemoryPool::allocate
     */
    void* allocate(size_t size, size_t alignment = AlignedAllocator::default_alignment()) override {
        if (size == 0) return nullptr;
        SizeClass* size_class = findSizeClass(size, alignment);
        if (!size_class) {
            void* p = AlignedAllocator::allocate(alignment, size);
            if (p) updateStatsOnAllocate(size);
            return p;
        }
        std::lock_guard<std::mutex> lk(size_class->mtx);
        if (size_class->free_list.empty()) {
            allocateNewPage(*size_class);
        }
        if (size_class->free_list.empty()) return nullptr;
        void* p = size_class->free_list.back();
        size_class->free_list.pop_back();
        updateStatsOnAllocate(size_class->size);
        return p;
    }
    
    /**
     * @brief 释放内存
     * @param ptr 要释放的内存指针
     * @param size 内存大小
     * @param alignment 内存对齐要求
     */
    void deallocate(void* ptr, size_t size, size_t alignment = AlignedAllocator::default_alignment()) override {
        if (!ptr) return;
        SizeClass* size_class = findSizeClass(size, alignment);
        if (!size_class) {
            AlignedAllocator::deallocate(ptr);
            updateStatsOnDeallocate(size);
            return;
        }
        std::lock_guard<std::mutex> lk(size_class->mtx);
        size_class->free_list.push_back(ptr);
        updateStatsOnDeallocate(size_class->size);
    }
    
    /**
     * @brief 获取支持的大小类数量
     * @return 大小类数量
     */
    size_t getSizeClassCount() const {
        std::lock_guard<std::mutex> lock(mutex_);
        return size_classes_.size();
    }
    
    /**
     * @brief 获取指定大小的建议对齐值
     * @param size 内存大小
     * @return 建议的对齐值
     */
    size_t getRecommendedAlignment(size_t size) const {
        return std::max(size_t(8), size);
    }

    // IMemoryPool 维护接口
    void shrink() override {
        // 策略：当某个size class所有块均空闲时，保留一页，释放多余页面
        for (auto& sc : size_classes_) {
            std::lock_guard<std::mutex> lk(sc.mtx);
            size_t total_blocks = 0;
            for (auto& pg : sc.pages) total_blocks += pg->blocks;
            if (total_blocks == 0) continue;
            if (sc.free_list.size() == total_blocks && sc.pages.size() > 1) {
                // 释放到仅保留一页
                auto keep = std::move(sc.pages.front());
                sc.pages.clear();
                sc.pages.push_back(std::move(keep));
                rebuildFreeList(sc);
            }
        }
    }

    void clear() override {
        std::lock_guard<std::mutex> lock(mutex_);
        for (auto& sc : size_classes_) {
            std::lock_guard<std::mutex> lk(sc.mtx);
            sc.pages.clear();
            sc.free_list.clear();
        }
        stats_ = {};
    }

    IMemoryPool::Statistics getStatistics() const override {
        std::lock_guard<std::mutex> lock(mutex_);
        return stats_;
    }

private:
    std::vector<SizeClass> size_classes_;
    mutable std::mutex mutex_;
    IMemoryPool::Statistics stats_{};
    
    void initializeSizeClasses() {
        // 常用大小：8, 16, 32, 64, 128, 256, 512, 1024 bytes（避免 vector 触发移动/拷贝构造问题）
        const size_t sizes[] = {8, 16, 32, 64, 128, 256, 512, 1024};
        size_classes_.reserve(sizeof(sizes) / sizeof(sizes[0]));
        for (size_t s : sizes) {
            size_classes_.push_back(SizeClass{s, getRecommendedAlignment(s)});
        }
    }
    
    SizeClass* findSizeClass(size_t size, size_t alignment) {
        std::lock_guard<std::mutex> lock(mutex_);
        for (auto& sc : size_classes_) {
            if (sc.size >= size && sc.alignment >= alignment) {
                return &sc;
            }
        }
        return nullptr;
    }
    
    void allocateNewPage(SizeClass& sc) {
        constexpr size_t target_bytes = 64 * 1024; // 64KB 页面
        size_t blocks = std::max<size_t>(16, target_bytes / std::max<size_t>(sc.stride, 1));
        size_t bytes = sc.stride * blocks;
        auto page = std::make_unique<Page>();
        page->buffer = AlignedAllocator::allocate(sc.alignment, bytes);
        if (!page->buffer) return;
        page->stride = sc.stride;
        page->blocks = blocks;
        // 切分为块并加入空闲表
        char* base = static_cast<char*>(page->buffer);
        for (size_t i = 0; i < blocks; ++i) {
            sc.free_list.push_back(base + i * sc.stride);
        }
        sc.pages.push_back(std::move(page));
    }

    void rebuildFreeList(SizeClass& sc) {
        sc.free_list.clear();
        for (auto& pg : sc.pages) {
            char* base = static_cast<char*>(pg->buffer);
            for (size_t i = 0; i < pg->blocks; ++i) {
                sc.free_list.push_back(base + i * pg->stride);
            }
        }
    }

    void updateStatsOnAllocate(std::size_t size) {
        std::lock_guard<std::mutex> lock(mutex_);
        stats_.total_allocations += 1;
        stats_.current_usage += size;
        if (stats_.current_usage > stats_.peak_usage) stats_.peak_usage = stats_.current_usage;
    }

    void updateStatsOnDeallocate(std::size_t size) {
        std::lock_guard<std::mutex> lock(mutex_);
        stats_.total_deallocations += 1;
        if (stats_.current_usage >= size) stats_.current_usage -= size; else stats_.current_usage = 0;
    }
};

}} // namespace fastexcel::memory
