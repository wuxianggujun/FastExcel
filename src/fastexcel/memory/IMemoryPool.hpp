/**
 * @file IMemoryPool.hpp
 * @brief 内存池统一接口定义（原始内存视角）
 */

#pragma once

#include <cstddef>

namespace fastexcel {
namespace memory {

/**
 * @brief 统一的内存池接口（按字节分配）。
 *
 * 说明：
 * - allocate/deallocate 面向“原始内存”语义，不负责对象构造/析构。
 * - 若需要对象生命周期管理，可使用本接口的默认辅助函数 construct/destroy。
 */
class IMemoryPool {
public:
    virtual ~IMemoryPool() = default;

    /**
     * @brief 分配指定大小与对齐的原始内存（不构造对象）。
     * @param size 字节数
     * @param alignment 对齐（默认使用平台推荐对齐）
     * @return 成功返回非空指针，失败返回nullptr
     */
    virtual void* allocate(std::size_t size, std::size_t alignment = alignof(std::max_align_t)) = 0;

    /**
     * @brief 释放由本池分配的原始内存（不调用析构）。
     * @param ptr 指针
     * @param size 字节数（部分实现可忽略）
     * @param alignment 对齐（部分实现可忽略）
     */
    virtual void deallocate(void* ptr, std::size_t size, std::size_t alignment = alignof(std::max_align_t)) = 0;

    /**
     * @brief 尝试收缩池，释放未使用内存（实现可按策略选择）。
     */
    virtual void shrink() = 0;

    /**
     * @brief 清空池中资源。
     */
    virtual void clear() = 0;

    /**
     * @brief 统计信息（统一字段，单位字节）。
     */
    struct Statistics {
        std::size_t current_usage = 0;       // 当前使用中的字节数（估算）
        std::size_t peak_usage = 0;          // 峰值使用字节数（估算）
        std::size_t total_allocations = 0;   // 分配次数（或累计字节，视实现定义）
        std::size_t total_deallocations = 0; // 释放次数（或累计字节，视实现定义）
    };

    /**
     * @brief 获取统计信息。
     */
    virtual Statistics getStatistics() const = 0;

    // ---------- 便捷模板：对象构造/析构（基于原始接口） ----------

    template <typename T, typename... Args>
    T* construct(Args&&... args) {
        void* mem = allocate(sizeof(T), alignof(T));
        if (!mem) return nullptr;
        return new (mem) T(static_cast<Args&&>(args)...);
    }

    template <typename T>
    void destroy(T* obj) {
        if (!obj) return;
        obj->~T();
        deallocate(static_cast<void*>(obj), sizeof(T), alignof(T));
    }
};

}} // namespace fastexcel::memory

