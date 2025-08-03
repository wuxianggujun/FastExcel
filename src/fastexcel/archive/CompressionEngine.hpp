#pragma once

#include <memory>
#include <string>
#include <vector>
#include <cstdint>

namespace fastexcel {
namespace archive {

/**
 * @brief 压缩引擎抽象接口
 * 
 * 提供统一的压缩接口，支持多种压缩后端（zlib, libdeflate等）
 */
class CompressionEngine {
public:
    /**
     * @brief 支持的压缩后端
     */
    enum class Backend {
        ZLIB,       // 标准 zlib 实现
        LIBDEFLATE  // 高性能 libdeflate 实现
    };
    
    /**
     * @brief 压缩结果结构
     */
    struct CompressionResult {
        bool success = false;
        size_t compressed_size = 0;
        std::string error_message;
    };
    
    /**
     * @brief 创建压缩引擎实例
     * @param backend 压缩后端类型
     * @param compression_level 压缩级别 (1-9)
     * @return 压缩引擎实例
     */
    static std::unique_ptr<CompressionEngine> create(Backend backend, int compression_level = 6);
    
    /**
     * @brief 获取可用的后端列表
     * @return 支持的后端列表
     */
    static std::vector<Backend> getAvailableBackends();
    
    /**
     * @brief 后端名称转换
     */
    static std::string backendToString(Backend backend);
    static Backend stringToBackend(const std::string& name);
    
    virtual ~CompressionEngine() = default;
    
    /**
     * @brief 压缩数据
     * @param input 输入数据
     * @param input_size 输入数据大小
     * @param output 输出缓冲区
     * @param output_capacity 输出缓冲区容量
     * @return 压缩结果
     */
    virtual CompressionResult compress(const void* input, size_t input_size,
                                     void* output, size_t output_capacity) = 0;
    
    /**
     * @brief 重置压缩器状态（用于重用）
     */
    virtual void reset() = 0;
    
    /**
     * @brief 获取引擎名称
     */
    virtual const char* name() const = 0;
    
    /**
     * @brief 获取当前压缩级别
     */
    virtual int getCompressionLevel() const = 0;
    
    /**
     * @brief 设置压缩级别
     * @param level 压缩级别 (1-9)
     * @return 是否成功
     */
    virtual bool setCompressionLevel(int level) = 0;
    
    /**
     * @brief 估算压缩后的最大大小
     * @param input_size 输入数据大小
     * @return 估算的最大压缩大小
     */
    virtual size_t getMaxCompressedSize(size_t input_size) const = 0;
    
    /**
     * @brief 获取引擎统计信息
     */
    struct Statistics {
        size_t total_input_bytes = 0;
        size_t total_output_bytes = 0;
        size_t compression_count = 0;
        double total_time_ms = 0.0;
        
        double getCompressionRatio() const {
            return total_input_bytes > 0 ? 
                static_cast<double>(total_output_bytes) / total_input_bytes : 0.0;
        }
        
        double getAverageSpeed() const {
            return total_time_ms > 0 ? 
                (total_input_bytes / 1024.0 / 1024.0) / (total_time_ms / 1000.0) : 0.0;
        }
    };
    
    virtual Statistics getStatistics() const = 0;
    virtual void resetStatistics() = 0;

protected:
    CompressionEngine() = default;
};

/**
 * @brief 自动选择最优压缩引擎
 * @param input_size 输入数据大小
 * @param compression_level 压缩级别
 * @return 推荐的后端类型
 */
CompressionEngine::Backend selectOptimalBackend(size_t input_size, int compression_level);

} // namespace archive
} // namespace fastexcel