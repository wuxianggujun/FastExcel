#pragma once

#include <vector>
#include <queue>
#include <memory>
#include <thread>
#include <mutex>
#include <condition_variable>
#include <future>
#include <functional>
#include "Exception.hpp"

namespace fastexcel {
namespace core {

/**
 * @brief 高性能线程池实现
 * 
 * 提供异步任务执行能力，支持任意可调用对象的并行执行
 */
class ThreadPool {
public:
    /**
     * @brief 构造函数
     * @param threads 线程数量，默认为硬件并发数
     */
    explicit ThreadPool(size_t threads = std::thread::hardware_concurrency());
    
    /**
     * @brief 析构函数，等待所有任务完成并清理资源
     */
    ~ThreadPool();
    
    /**
     * @brief 提交任务到线程池
     * @tparam F 函数类型
     * @tparam Args 参数类型
     * @param f 要执行的函数
     * @param args 函数参数
     * @return std::future 用于获取任务结果
     */
    template<class F, class... Args>
    auto enqueue(F&& f, Args&&... args)
        -> std::future<typename std::invoke_result<F, Args...>::type>;
    
    /**
     * @brief 获取线程池大小
     * @return 线程数量
     */
    size_t size() const { return workers_.size(); }
    
    /**
     * @brief 获取待处理任务数量
     * @return 任务队列大小
     */
    size_t pending_tasks() const {
        std::unique_lock<std::mutex> lock(queue_mutex_);
        return tasks_.size();
    }
    
    /**
     * @brief 等待所有任务完成
     */
    void wait_for_all_tasks();

private:
    // 工作线程
    std::vector<std::thread> workers_;
    
    // 任务队列
    std::queue<std::function<void()>> tasks_;
    
    // 同步原语
    mutable std::mutex queue_mutex_;
    std::condition_variable condition_;
    std::condition_variable finished_;
    
    // 控制标志
    bool stop_;
    size_t active_tasks_;
};

// 模板方法实现
template<class F, class... Args>
auto ThreadPool::enqueue(F&& f, Args&&... args)
    -> std::future<typename std::invoke_result<F, Args...>::type> {
    
    using return_type = typename std::invoke_result<F, Args...>::type;
    
    auto task = std::make_shared<std::packaged_task<return_type()>>(
        std::bind(std::forward<F>(f), std::forward<Args>(args)...)
    );
    
    std::future<return_type> res = task->get_future();
    
    {
        std::unique_lock<std::mutex> lock(queue_mutex_);
        
        // 不允许在停止状态下添加任务
        if (stop_) {
            FASTEXCEL_THROW_OP("enqueue on stopped ThreadPool");
        }
        
        tasks_.emplace([task]() { (*task)(); });
        ++active_tasks_;
    }
    
    condition_.notify_one();
    return res;
}

}} // namespace fastexcel::core