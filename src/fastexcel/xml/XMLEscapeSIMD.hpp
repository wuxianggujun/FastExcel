#pragma once

#include <cstddef>
#include <functional>
#include <string>

namespace fastexcel {
namespace xml {

/**
 * @brief SIMD优化的XML转义处理器
 * 
 * 使用Highway库进行跨平台SIMD优化，加速XML字符转义处理
 */
class XMLEscapeSIMD {
public:
    // 写入回调类型
    using WriteCallback = std::function<void(const std::string&)>;
    
    /**
     * @brief 初始化SIMD转义器并检测CPU特性
     */
    static void initialize();
    
    /**
     * @brief 使用SIMD优化进行属性值转义
     * @param text 要转义的文本
     * @param length 文本长度
     * @param writer 写入回调函数
     */
    static void escapeAttributesSIMD(const char* text, size_t length, WriteCallback writer);
    
    /**
     * @brief 使用SIMD优化进行文本内容转义
     * @param text 要转义的文本
     * @param length 文本长度
     * @param writer 写入回调函数
     */
    static void escapeDataSIMD(const char* text, size_t length, WriteCallback writer);
    
    /**
     * @brief 检查当前平台是否支持SIMD优化
     * @return true 如果支持SIMD优化
     */
    static bool isSIMDSupported();
    
private:
    static bool simd_initialized_;
    static bool simd_supported_;
    
    /**
     * @brief 标量版本的属性转义实现（SIMD回退版本）
     */
    static void escapeAttributesScalar(const char* text, size_t length, WriteCallback writer);
    
    /**
     * @brief 标量版本的数据转义实现（SIMD回退版本）
     */
    static void escapeDataScalar(const char* text, size_t length, WriteCallback writer);
};

}} // namespace fastexcel::xml