#include "fastexcel/utils/Logger.hpp"
#include "LibDeflateEngine.hpp"

#ifdef FASTEXCEL_HAS_LIBDEFLATE

#include <stdexcept>
#include <algorithm>
#include <cstring>

namespace fastexcel {
namespace archive {

LibDeflateEngine::LibDeflateEngine(int compression_level)
    : compressor_(nullptr)
    , compression_level_(std::clamp(compression_level, 1, 12))  // libdeflate 支持 1-12
    , stats_{} {
    
    auto result = initializeCompressor();
    if (result.hasError()) {
        // 无法在构造函数中返回Result，但我们可以记录错误状态
        // 或者保留异常作为最后手段
        compressor_ = nullptr;
    }
}

LibDeflateEngine::~LibDeflateEngine() {
    cleanupCompressor();
}

VoidResult LibDeflateEngine::initializeCompressor() {
    if (compressor_) {
        cleanupCompressor();
    }
    
    if (!isValidCompressionLevel(compression_level_)) {
        return makeError(ErrorCode::InvalidArgument, "Invalid compression level");
    }
    
    compressor_ = libdeflate_alloc_compressor(compression_level_);
    if (!compressor_) {
        return makeError(ErrorCode::OutOfMemory, "Failed to allocate libdeflate compressor");
    }
    return success();
}

void LibDeflateEngine::cleanupCompressor() {
    if (compressor_) {
        libdeflate_free_compressor(compressor_);
        compressor_ = nullptr;
    }
}

Result<size_t> LibDeflateEngine::compress(
    const void* input, size_t input_size,
    void* output, size_t output_capacity) {
    
    if (!compressor_) {
        return makeError(ErrorCode::InternalError, "Compressor not initialized");
    }
    
    if (!input || input_size == 0 || !output || output_capacity == 0) {
        return makeError(ErrorCode::InvalidArgument, "Invalid input parameters");
    }
    
    auto start_time = std::chrono::high_resolution_clock::now();
    
    // 使用 libdeflate 进行原始 DEFLATE 压缩
    // 这与 zlib 的 raw deflate 输出兼容
    size_t compressed_size = libdeflate_deflate_compress(
        compressor_,
        input, input_size,
        output, output_capacity
    );
    
    auto end_time = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);
    double time_ms = duration.count() / 1000.0;
    
    if (compressed_size == 0) {
        return makeError(ErrorCode::ZipCompressionError, "Compression failed - output buffer too small or compression error");
    }
    
    // 更新统计信息
    updateStatistics(input_size, compressed_size, time_ms);
    
    return makeExpected(compressed_size);
}

VoidResult LibDeflateEngine::reset() {
    // libdeflate 的压缩器是无状态的，不需要重置
    // 每次调用 libdeflate_deflate_compress 都是独立的
    return success();
}

VoidResult LibDeflateEngine::setCompressionLevel(int level) {
    if (!isValidCompressionLevel(level)) {
        return makeError(ErrorCode::InvalidArgument, "Invalid compression level");
    }
    
    if (level == compression_level_) {
        return success();
    }
    
    compression_level_ = level;
    
    // 重新初始化压缩器以应用新的压缩级别
    return initializeCompressor();
}

size_t LibDeflateEngine::getMaxCompressedSize(size_t input_size) const {
    if (!compressor_) {
        // 使用保守估算
        return input_size + (input_size >> 8) + 64;
    }
    
    // 使用 libdeflate 的精确边界计算
    return libdeflate_deflate_compress_bound(compressor_, input_size);
}

VoidResult LibDeflateEngine::resetStatistics() {
    stats_ = Statistics{};
    return success();
}

void LibDeflateEngine::updateStatistics(size_t input_size, size_t output_size, double time_ms) {
    stats_.total_input_bytes += input_size;
    stats_.total_output_bytes += output_size;
    stats_.compression_count++;
    stats_.total_time_ms += time_ms;
}

bool LibDeflateEngine::isValidCompressionLevel(int level) {
    // libdeflate 支持压缩级别 0-12
    // 0 = 无压缩，1 = 最快，6 = 默认，12 = 最慢/最好压缩
    return level >= 0 && level <= 12;
}

} // namespace archive
} // namespace fastexcel

#endif // FASTEXCEL_HAS_LIBDEFLATE
