/**
 * @file UserLayerExample.cpp
 * @brief 展示如何在用户层API中使用异常转换层
 */

#include "fastexcel/core/ExceptionBridge.hpp"
#include "fastexcel/core/MemoryPool.hpp"
#include "fastexcel/archive/CompressionEngine.hpp"
#include <iostream>
#include <memory>

namespace fastexcel {
namespace core {

/**
 * @brief 用户层Workbook类示例
 * 
 * 展示如何在用户友好的API中使用底层的Result/Expected类型
 */
class UserWorkbook {
private:
    std::unique_ptr<MemoryPool> memory_pool_;
    std::unique_ptr<archive::CompressionEngine> compression_engine_;

public:
    /**
     * @brief 构造函数 - 用户层API，抛出异常
     */
    UserWorkbook() {
        // 直接创建内存池（MemoryPool构造函数已经处理了Result）
        memory_pool_ = std::make_unique<MemoryPool>(1024, 16);
        
        // 底层使用Result，用户层自动转换为异常
        auto engine_result = archive::CompressionEngine::create(
            archive::CompressionEngine::Backend::ZLIB);
        compression_engine_ = FASTEXCEL_UNWRAP(std::move(engine_result));
    }
    
    /**
     * @brief 分配内存 - 用户友好的异常接口
     * @param size 内存大小
     * @return 内存指针
     * @throws MemoryException 如果分配失败
     */
    void* allocateMemory(size_t size) {
        auto result = memory_pool_->allocate(size);
        return FASTEXCEL_UNWRAP(std::move(result));
    }
    
    /**
     * @brief 释放内存 - 用户友好的异常接口
     * @param ptr 内存指针
     * @throws ParameterException 如果指针无效
     */
    void deallocateMemory(void* ptr) {
        auto result = memory_pool_->deallocate(ptr);
        FASTEXCEL_UNWRAP(std::move(result)); // void返回类型
    }
    
    /**
     * @brief 压缩数据 - 用户友好的异常接口
     * @param input 输入数据
     * @param input_size 输入大小
     * @param output 输出缓冲区
     * @param output_capacity 输出容量
     * @return 压缩后的大小
     * @throws CompressionException 如果压缩失败
     */
    size_t compressData(const void* input, size_t input_size,
                       void* output, size_t output_capacity) {
        auto result = compression_engine_->compress(input, input_size, output, output_capacity);
        return FASTEXCEL_UNWRAP(std::move(result));
    }
    
    /**
     * @brief 高级API：提供可选的错误处理方式
     */
    class SafeOperations {
    private:
        UserWorkbook& workbook_;
        
    public:
        explicit SafeOperations(UserWorkbook& wb) : workbook_(wb) {}
        
        /**
         * @brief 安全分配内存 - 不抛出异常的版本
         * @param size 内存大小
         * @return 包装的结果，可以选择获取值或检查错误
         */
        UserAPIWrapper<void*> tryAllocateMemory(size_t size) {
            auto result = workbook_.memory_pool_->allocate(size);
            return wrapForUser(std::move(result));
        }
        
        /**
         * @brief 安全释放内存 - 不抛出异常的版本
         * @param ptr 内存指针
         * @return 包装的结果，可以检查是否成功
         */
        VoidUserAPIWrapper tryDeallocateMemory(void* ptr) {
            auto result = workbook_.memory_pool_->deallocate(ptr);
            return wrapForUser(std::move(result));
        }
    };
    
    /**
     * @brief 获取安全操作接口
     */
    SafeOperations safe() { return SafeOperations(*this); }
};

} // namespace core
} // namespace fastexcel

/**
 * @brief 使用示例
 */
void demonstrateUsage() {
    using namespace fastexcel::core;
    
    try {
        // 1. 传统异常风格的用户API
        UserWorkbook workbook;
        
        // 自动转换：底层Result -> 用户层异常
        void* memory = workbook.allocateMemory(1024);
        
        // 底层错误会自动转换为对应的异常类型
        size_t compressed_size = workbook.compressData(
            "Hello World", 11, memory, 1024);
            
        workbook.deallocateMemory(memory);
        
        std::cout << "压缩成功，大小: " << compressed_size << std::endl;
        
    } catch (const MemoryException& e) {
        std::cerr << "内存错误: " << e.what() << std::endl;
    } catch (const OperationException& e) {
        std::cerr << "操作错误: " << e.what() << std::endl;
    } catch (const FastExcelException& e) {
        std::cerr << "一般错误: " << e.what() << std::endl;
    }
    
    // 2. 可选的安全API（不抛异常）
    UserWorkbook workbook;
    auto safe_ops = workbook.safe();
    
    // 方式A: 检查成功/失败但不获取详细错误
    auto memory_result = safe_ops.tryAllocateMemory(1024);
    if (memory_result.isSuccess()) {
        void* memory = memory_result.get(); // 这里仍会抛异常，但我们已经检查过了
        
        auto dealloc_result = safe_ops.tryDeallocateMemory(memory);
        if (!dealloc_result.isSuccess()) {
            std::cerr << "释放内存失败: " << dealloc_result.getErrorMessage() << std::endl;
        }
    } else {
        std::cerr << "分配内存失败: " << memory_result.getErrorMessage() << std::endl;
    }
    
    std::cout << "演示完成！" << std::endl;
}