/**
 * @file IObjectPool.hpp
 * @brief 统一的“类型化对象内存池”接口定义
 */

#pragma once

#include <cstddef>

namespace fastexcel {
namespace memory {

/**
 * @brief 类型T对象的统一内存池接口。
 *
 * 约定：
 * - allocate() 默认构造一个 T 并返回指针；
 * - deallocate() 负责调用析构并回收对象；
 * - 统计字段单位为“对象个数”或“估算值”，由实现解释。
 */
template <typename T>
class IObjectPool {
public:
    virtual ~IObjectPool() = default;

    // 分配一个默认构造的对象
    virtual T* allocate() = 0;
    // 回收对象（应调用析构）
    virtual void deallocate(T* obj) = 0;

    // 统计与维护
    virtual std::size_t getCurrentUsage() const noexcept = 0;
    virtual std::size_t getPeakUsage() const noexcept = 0;
    virtual std::size_t getTotalAllocated() const noexcept = 0;
    virtual void shrink() = 0;
    virtual void clear() = 0;
};

}} // namespace fastexcel::memory

