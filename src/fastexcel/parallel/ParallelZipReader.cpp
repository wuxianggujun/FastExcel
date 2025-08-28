#include "fastexcel/parallel/ParallelZipReader.hpp"

#include "fastexcel/archive/ZipError.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <algorithm>
#include <thread>

namespace fastexcel {
namespace parallel {

// ParallelZipReader implementation
ParallelZipReader::ParallelZipReader(const core::Path& zip_path, const Config& config)
    : config_(config), zip_path_(zip_path) {
    
    // Initialize thread pool
    thread_pool_ = std::make_unique<core::ThreadPool>(config_.thread_count);
    
    // Initialize reader pool
    reader_pool_ = std::make_unique<ReaderPool>(zip_path_, config_.thread_count);
    
    // Initialize cache if enabled
    if (config_.enable_cache) {
        cache_ = std::make_unique<Cache>(config_.cache_size_limit);
    }
    
    // Start worker threads
    stop_flag_.store(false);
    for (size_t i = 0; i < config_.thread_count; ++i) {
        worker_threads_.emplace_back(&ParallelZipReader::workerThread, this);
    }
    
    FASTEXCEL_LOG_DEBUG("ParallelZipReader initialized with {} threads", config_.thread_count);
}

ParallelZipReader::~ParallelZipReader() {
    stop_flag_.store(true);
    queue_cv_.notify_all();
    
    for (auto& thread : worker_threads_) {
        if (thread.joinable()) {
            thread.join();
        }
    }
}

std::unordered_map<std::string, std::vector<uint8_t>> 
ParallelZipReader::extractFilesParallel(const std::vector<std::string>& paths) {
    std::unordered_map<std::string, std::vector<uint8_t>> results;
    std::vector<std::future<std::vector<uint8_t>>> futures;
    
    // Submit all tasks
    for (const auto& path : paths) {
        futures.push_back(extractFileAsync(path));
    }
    
    // Collect results
    for (size_t i = 0; i < paths.size(); ++i) {
        try {
            results[paths[i]] = futures[i].get();
        } catch (const std::exception& e) {
            FASTEXCEL_LOG_ERROR("Failed to extract file {}: {}", paths[i], e.what());
        }
    }
    
    return results;
}

std::future<std::vector<uint8_t>> 
ParallelZipReader::extractFileAsync(const std::string& path) {
    auto task = std::make_shared<ExtractionTask>(path);
    
    {
        std::lock_guard<std::mutex> lock(queue_mutex_);
        task_queue_.push(task);
    }
    queue_cv_.notify_one();
    
    return std::move(task->future);
}

void ParallelZipReader::prefetchFiles(const std::vector<std::string>& paths) {
    for (const auto& path : paths) {
        thread_pool_->enqueue([this, path]() {
            try {
                extractFileInternal(path);
            } catch (const std::exception& e) {
                FASTEXCEL_LOG_WARN("Failed to prefetch file {}: {}", path, e.what());
            }
        });
    }
}

void ParallelZipReader::processFilesInParallel(
    const std::vector<std::string>& paths,
    std::function<void(const std::string&, const std::vector<uint8_t>&)> processor) {
    
    std::vector<std::future<void>> futures;
    
    for (const auto& path : paths) {
        auto future = thread_pool_->enqueue([this, path, processor]() {
            try {
                auto data = extractFileInternal(path);
                processor(path, data);
            } catch (const std::exception& e) {
                FASTEXCEL_LOG_ERROR("Failed to process file {}: {}", path, e.what());
            }
        });
        
        futures.push_back(std::move(future));
    }
    
    // Wait for all tasks to complete
    for (auto& future : futures) {
        future.wait();
    }
}

void ParallelZipReader::workerThread() {
    while (!stop_flag_.load()) {
        std::shared_ptr<ExtractionTask> task;
        
        // Get task from queue
        {
            std::unique_lock<std::mutex> lock(queue_mutex_);
            queue_cv_.wait(lock, [this] { return !task_queue_.empty() || stop_flag_.load(); });
            
            if (stop_flag_.load()) {
                break;
            }
            
            if (!task_queue_.empty()) {
                task = task_queue_.front();
                task_queue_.pop();
            }
        }
        
        if (task) {
            try {
                auto result = extractFileInternal(task->path);
                task->promise.set_value(std::move(result));
            } catch (const std::exception&) {
                task->promise.set_exception(std::current_exception());
            }
        }
    }
}

std::vector<uint8_t> ParallelZipReader::extractFileInternal(const std::string& path) {
    // Check cache first
    if (cache_) {
        std::vector<uint8_t> cached_result;
        if (cache_->get(path, cached_result)) {
            cache_hits_.fetch_add(1);
            return cached_result;
        }
        cache_misses_.fetch_add(1);
    }
    
    // Extract from ZIP
    auto reader = reader_pool_->acquire();
    if (!reader) {
        throw std::runtime_error("Failed to acquire ZIP reader");
    }
    
    try {
        std::vector<uint8_t> result;
        auto error = reader->extractFile(path, result);
        if (error != archive::ZipError::Ok) {
            reader_pool_->release(reader);
            throw std::runtime_error("Failed to extract file: " + path);
        }
        
        reader_pool_->release(reader);
        
        // Cache the result
        if (cache_ && !result.empty()) {
            cache_->put(path, result);
        }
        
        return result;
    } catch (...) {
        reader_pool_->release(reader);
        throw;
    }
}

void ParallelZipReader::clearCache() {
    if (cache_) {
        cache_->clear();
    }
}

ParallelZipReader::CacheStats ParallelZipReader::getCacheStats() const {
    CacheStats stats;
    stats.hit_count = cache_hits_.load();
    stats.miss_count = cache_misses_.load();
    stats.cache_size = cache_ ? cache_->size() : 0;
    
    size_t total = stats.hit_count + stats.miss_count;
    stats.hit_rate = total > 0 ? static_cast<double>(stats.hit_count) / total : 0.0;
    
    return stats;
}

// ReaderPool implementation
ParallelZipReader::ReaderPool::ReaderPool(const core::Path& path, size_t pool_size) 
    : zip_path_(path) {
    
    for (size_t i = 0; i < pool_size; ++i) {
        auto reader = std::make_shared<archive::ZipReader>(zip_path_);
        if (reader->open()) {
            readers_.push_back(reader);
            available_.push(reader);
        }
    }
    
    if (available_.empty()) {
        throw std::runtime_error("Failed to initialize any ZIP readers");
    }
}

ParallelZipReader::ReaderPool::~ReaderPool() {
    std::lock_guard<std::mutex> lock(mutex_);
    readers_.clear();
    // Clear the queue
    while (!available_.empty()) {
        available_.pop();
    }
}

std::shared_ptr<archive::ZipReader> ParallelZipReader::ReaderPool::acquire() {
    std::unique_lock<std::mutex> lock(mutex_);
    cv_.wait(lock, [this] { return !available_.empty(); });
    
    auto reader = available_.front();
    available_.pop();
    return reader;
}

void ParallelZipReader::ReaderPool::release(std::shared_ptr<archive::ZipReader> reader) {
    std::lock_guard<std::mutex> lock(mutex_);
    available_.push(reader);
    cv_.notify_one();
}

// Cache implementation
ParallelZipReader::Cache::Cache(size_t size_limit) 
    : size_limit_(size_limit), current_size_(0) {
}

bool ParallelZipReader::Cache::get(const std::string& key, std::vector<uint8_t>& value) {
    std::shared_lock<std::shared_mutex> lock(mutex_);
    
    auto it = cache_.find(key);
    if (it != cache_.end()) {
        value = it->second.data;
        it->second.last_access = std::chrono::steady_clock::now();
        return true;
    }
    
    return false;
}

void ParallelZipReader::Cache::put(const std::string& key, const std::vector<uint8_t>& value) {
    std::unique_lock<std::shared_mutex> lock(mutex_);
    
    // Check if we need to evict
    size_t value_size = value.size();
    while (current_size_ + value_size > size_limit_ && !cache_.empty()) {
        evictLRU();
    }
    
    // Don't cache if value is too large
    if (value_size > size_limit_) {
        return;
    }
    
    // Insert new entry
    CacheEntry entry;
    entry.data = value;
    entry.last_access = std::chrono::steady_clock::now();
    
    // Remove existing entry if present
    auto existing = cache_.find(key);
    if (existing != cache_.end()) {
        current_size_ -= existing->second.data.size();
        cache_.erase(existing);
    }
    
    cache_[key] = std::move(entry);
    current_size_ += value_size;
}

void ParallelZipReader::Cache::clear() {
    std::unique_lock<std::shared_mutex> lock(mutex_);
    cache_.clear();
    current_size_ = 0;
}

size_t ParallelZipReader::Cache::size() const {
    std::shared_lock<std::shared_mutex> lock(mutex_);
    return current_size_;
}

void ParallelZipReader::Cache::evictLRU() {
    if (cache_.empty()) return;
    
    auto oldest_it = cache_.begin();
    auto oldest_time = oldest_it->second.last_access;
    
    for (auto it = cache_.begin(); it != cache_.end(); ++it) {
        if (it->second.last_access < oldest_time) {
            oldest_it = it;
            oldest_time = it->second.last_access;
        }
    }
    
    current_size_ -= oldest_it->second.data.size();
    cache_.erase(oldest_it);
}

// ParallelWorksheetLoader implementation
std::vector<ParallelWorksheetLoader::WorksheetData> 
ParallelWorksheetLoader::loadWorksheetsParallel(
    ParallelZipReader& zip_reader,
    const std::vector<std::pair<std::string, std::string>>& worksheet_paths) {
    
    std::vector<std::string> paths;
    paths.reserve(worksheet_paths.size());
    for (const auto& pair : worksheet_paths) {
        paths.push_back(pair.second);
    }
    
    auto files_data = zip_reader.extractFilesParallel(paths);
    std::vector<WorksheetData> results;
    results.reserve(worksheet_paths.size());
    
    for (size_t i = 0; i < worksheet_paths.size(); ++i) {
        const auto& [name, path] = worksheet_paths[i];
        auto it = files_data.find(path);
        if (it != files_data.end()) {
            WorksheetData data;
            data.name = name;
            data.path = path;
            data.content = std::move(it->second);
            data.row_count = 0; // Will be filled by parser
            data.col_count = 0; // Will be filled by parser
            results.push_back(std::move(data));
        }
    }
    
    return results;
}

void ParallelWorksheetLoader::processWorksheetsPipeline(
    ParallelZipReader& zip_reader,
    const std::vector<std::string>& worksheet_paths,
    std::function<void(const WorksheetData&)> processor) {
    
    zip_reader.processFilesInParallel(worksheet_paths, 
        [&processor](const std::string& path, const std::vector<uint8_t>& data) {
            WorksheetData worksheet_data;
            worksheet_data.name = path; // Extract name from path if needed
            worksheet_data.path = path;
            worksheet_data.content = data;
            worksheet_data.row_count = 0;
            worksheet_data.col_count = 0;
            
            processor(worksheet_data);
        });
}

}} // namespace fastexcel::parallel
