#include "fastexcel/utils/Logger.hpp"
#include "ZlibEngine.hpp"
#include <stdexcept>
#include <algorithm>
#include <cstring>

namespace fastexcel {
namespace archive {

ZlibEngine::ZlibEngine(int compression_level)
    : stream_(std::make_unique<z_stream>())
    , compression_level_(std::clamp(compression_level, 1, 9))
    , initialized_(false)
    , stats_{} {
    
    auto result = initializeStream();
    if (result.hasError()) {
        initialized_ = false;
        // 构造函数中无法返回Result，保留初始化失败的状态
    }
}

ZlibEngine::~ZlibEngine() {
    cleanupStream();
}

VoidResult ZlibEngine::initializeStream() {
    if (initialized_) {
        cleanupStream();
    }
    
    std::memset(stream_.get(), 0, sizeof(z_stream));
    
    // 初始化 deflate 流
    // 使用 -15 窗口位数生成原始 deflate 数据（无 zlib 头部）
    int ret = deflateInit2(stream_.get(), compression_level_, Z_DEFLATED, 
                          -15, 8, Z_DEFAULT_STRATEGY);
    
    if (ret != Z_OK) {
        return makeError(ErrorCode::ZipError, "Failed to initialize zlib deflate: " + std::to_string(ret));
    }
    
    initialized_ = true;
    return success();
}

void ZlibEngine::cleanupStream() {
    if (initialized_ && stream_) {
        deflateEnd(stream_.get());
        initialized_ = false;
    }
}

Result<size_t> ZlibEngine::compress(
    const void* input, size_t input_size,
    void* output, size_t output_capacity) {
    
    if (!initialized_) {
        return makeError(ErrorCode::InternalError, "Engine not initialized");
    }
    
    if (!input || input_size == 0 || !output || output_capacity == 0) {
        return makeError(ErrorCode::InvalidArgument, "Invalid input parameters");
    }
    
    auto start_time = std::chrono::high_resolution_clock::now();
    
    // 重置流状态以重用
    int ret = deflateReset(stream_.get());
    if (ret != Z_OK) {
        return makeError(ErrorCode::ZipError, "Failed to reset deflate stream");
    }
    
    // 设置输入和输出
    stream_->next_in = static_cast<Bytef*>(const_cast<void*>(input));
    stream_->avail_in = static_cast<uInt>(input_size);
    stream_->next_out = static_cast<Bytef*>(output);
    stream_->avail_out = static_cast<uInt>(output_capacity);
    
    // 执行压缩
    ret = deflate(stream_.get(), Z_FINISH);
    if (ret != Z_STREAM_END) {
        return makeError(ErrorCode::ZipError, "Deflate failed with code: " + std::to_string(ret));
    }
    
    // 获取压缩后的大小
    size_t compressed_size = output_capacity - stream_->avail_out;
    
    auto end_time = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);
    double time_ms = duration.count() / 1000.0;
    
    // 更新统计信息
    updateStatistics(input_size, compressed_size, time_ms);
    
    return makeExpected(compressed_size);
}

VoidResult ZlibEngine::reset() {
    if (initialized_) {
        deflateReset(stream_.get());
    }
    return success();
}

VoidResult ZlibEngine::setCompressionLevel(int level) {
    int new_level = std::clamp(level, 1, 9);
    if (new_level == compression_level_) {
        return success();
    }
    
    compression_level_ = new_level;
    
    // 重新初始化流以应用新的压缩级别
    return initializeStream();
}

size_t ZlibEngine::getMaxCompressedSize(size_t input_size) const {
    // 使用更精确的估算公式，避免 deflateBound 的过度分配
    // 基于经验：最坏情况下压缩后大小约为原始大小 + 0.1% + 64 字节
    return input_size + (input_size >> 8) + 64;
}

VoidResult ZlibEngine::resetStatistics() {
    stats_ = Statistics{};
    return success();
}

void ZlibEngine::updateStatistics(size_t input_size, size_t output_size, double time_ms) {
    stats_.total_input_bytes += input_size;
    stats_.total_output_bytes += output_size;
    stats_.compression_count++;
    stats_.total_time_ms += time_ms;
}

} // namespace archive
} // namespace fastexcel
