#pragma once

#include "CompressionEngine.hpp"

#ifdef FASTEXCEL_HAS_LIBDEFLATE
#include <libdeflate.h>
#include <memory>
#include <chrono>
#include <stdexcept>

namespace fastexcel {
namespace archive {

/**
 * @brief 基于 libdeflate 的高性能压缩引擎实现
 * 
 * libdeflate 是一个专门针对 DEFLATE 算法优化的库，
 * 相比 zlib 可以提供 1.8-2.5x 的性能提升
 */
class LibDeflateEngine : public CompressionEngine {
public:
    /**
     * @brief 构造函数
     * @param compression_level 压缩级别 (1-12, libdeflate 支持更高级别)
     */
    explicit LibDeflateEngine(int compression_level = 6);
    
    /**
     * @brief 析构函数
     */
    ~LibDeflateEngine() override;
    
    // CompressionEngine 接口实现
    CompressionResult compress(const void* input, size_t input_size,
                             void* output, size_t output_capacity) override;
    
    void reset() override;
    const char* name() const override { return "libdeflate"; }
    int getCompressionLevel() const override { return compression_level_; }
    bool setCompressionLevel(int level) override;
    size_t getMaxCompressedSize(size_t input_size) const override;
    Statistics getStatistics() const override { return stats_; }
    void resetStatistics() override;

private:
    struct libdeflate_compressor* compressor_;
    int compression_level_;
    Statistics stats_;
    
    /**
     * @brief 初始化 libdeflate 压缩器
     * @return 是否成功
     */
    bool initializeCompressor();
    
    /**
     * @brief 清理 libdeflate 压缩器
     */
    void cleanupCompressor();
    
    /**
     * @brief 更新统计信息
     */
    void updateStatistics(size_t input_size, size_t output_size, double time_ms);
    
    /**
     * @brief 验证压缩级别是否有效
     */
    static bool isValidCompressionLevel(int level);
};

} // namespace archive
} // namespace fastexcel

#else // !FASTEXCEL_HAS_LIBDEFLATE

#include <stdexcept>

namespace fastexcel {
namespace archive {

// 当 libdeflate 不可用时的占位符类
class LibDeflateEngine : public CompressionEngine {
public:
    explicit LibDeflateEngine(int) {
        throw std::runtime_error("LibDeflateEngine: libdeflate support not compiled in");
    }
    
    CompressionResult compress(const void*, size_t, void*, size_t) override {
        throw std::runtime_error("LibDeflateEngine: libdeflate support not compiled in");
    }
    
    void reset() override {}
    const char* name() const override { return "libdeflate (unavailable)"; }
    int getCompressionLevel() const override { return 0; }
    bool setCompressionLevel(int) override { return false; }
    size_t getMaxCompressedSize(size_t) const override { return 0; }
    Statistics getStatistics() const override { return Statistics{}; }
    void resetStatistics() override {}
};

} // namespace archive
} // namespace fastexcel

#endif // FASTEXCEL_HAS_LIBDEFLATE