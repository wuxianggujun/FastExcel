#pragma once

#include "CompressionEngine.hpp"
#include <zlib.h>
#include <memory>
#include <chrono>

namespace fastexcel {
namespace archive {

/**
 * @brief 基于 zlib 的压缩引擎实现
 */
class ZlibEngine : public CompressionEngine {
public:
    /**
     * @brief 构造函数
     * @param compression_level 压缩级别 (1-9)
     */
    explicit ZlibEngine(int compression_level = 6);
    
    /**
     * @brief 析构函数
     */
    ~ZlibEngine() override;
    
    // CompressionEngine 接口实现
    Result<size_t> compress(const void* input, size_t input_size,
                           void* output, size_t output_capacity) override;
    
    VoidResult reset() override;
    const char* name() const override { return "zlib"; }
    int getCompressionLevel() const override { return compression_level_; }
    VoidResult setCompressionLevel(int level) override;
    size_t getMaxCompressedSize(size_t input_size) const override;
    Statistics getStatistics() const override { return stats_; }
    VoidResult resetStatistics() override;

private:
    std::unique_ptr<z_stream> stream_;
    int compression_level_;
    bool initialized_;
    Statistics stats_;
    
    /**
     * @brief 初始化 zlib 流
     * @return VoidResult 成功或失败信息
     */
    VoidResult initializeStream();
    
    /**
     * @brief 清理 zlib 流
     */
    void cleanupStream();
    
    /**
     * @brief 更新统计信息
     */
    void updateStatistics(size_t input_size, size_t output_size, double time_ms);
};

} // namespace archive
} // namespace fastexcel