#include "fastexcel/utils/ThreadPool.hpp"
#include "fastexcel/utils/Logger.hpp"

namespace fastexcel {
namespace utils {

ThreadPool::ThreadPool(size_t threads) 
    : stop_(false), active_tasks_(0) {
    
    if (threads == 0) {
        threads = std::thread::hardware_concurrency();
        if (threads == 0) threads = 4; // 默认4个线程
    }
    
    LOG_INFO("Creating ThreadPool with {} threads", threads);
    
    workers_.reserve(threads);
    
    for (size_t i = 0; i < threads; ++i) {
        workers_.emplace_back([this] {
            for (;;) {
                std::function<void()> task;
                
                {
                    std::unique_lock<std::mutex> lock(this->queue_mutex_);
                    
                    // 等待任务或停止信号
                    this->condition_.wait(lock, [this] {
                        return this->stop_ || !this->tasks_.empty();
                    });
                    
                    // 如果停止且没有任务，退出
                    if (this->stop_ && this->tasks_.empty()) {
                        return;
                    }
                    
                    // 获取任务
                    task = std::move(this->tasks_.front());
                    this->tasks_.pop();
                }
                
                // 执行任务
                try {
                    task();
                } catch (const std::exception& e) {
                    LOG_ERROR("ThreadPool task exception: {}", e.what());
                } catch (...) {
                    LOG_ERROR("ThreadPool task unknown exception");
                }
                
                // 任务完成，更新计数器
                {
                    std::unique_lock<std::mutex> lock(this->queue_mutex_);
                    --this->active_tasks_;
                    if (this->active_tasks_ == 0 && this->tasks_.empty()) {
                        this->finished_.notify_all();
                    }
                }
            }
        });
    }
    
    LOG_DEBUG("ThreadPool created successfully with {} worker threads", workers_.size());
}

ThreadPool::~ThreadPool() {
    LOG_DEBUG("Destroying ThreadPool...");
    
    {
        std::unique_lock<std::mutex> lock(queue_mutex_);
        stop_ = true;
    }
    
    condition_.notify_all();
    
    for (std::thread &worker : workers_) {
        if (worker.joinable()) {
            worker.join();
        }
    }
    
    LOG_INFO("ThreadPool destroyed, {} threads joined", workers_.size());
}

void ThreadPool::wait_for_all_tasks() {
    std::unique_lock<std::mutex> lock(queue_mutex_);
    finished_.wait(lock, [this] {
        return this->tasks_.empty() && this->active_tasks_ == 0;
    });
}

}} // namespace fastexcel::utils