/**
 * @file minimal_destructor_test.cpp
 * @brief 最小化析构函数测试
 * 
 * 专门测试PoolAllocator析构时的问题
 */

#include "fastexcel/memory/PoolAllocator.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <iostream>
#include <vector>

using namespace fastexcel;
using namespace fastexcel::memory;

void testVectorAllocatorDestruction() {
    std::cout << "=== Testing Vector Allocator Destruction ===\n";
    
    {
        std::cout << "1. Creating PoolAllocator<std::vector<int>>...\n";
        PoolAllocator<std::vector<int>> vector_allocator;
        
        std::cout << "2. Allocating some vectors...\n";
        std::vector<std::vector<int>*> vectors;
        
        for (int i = 0; i < 3; ++i) {
            auto* vec = vector_allocator.allocate(1);
            vector_allocator.construct(vec);
            vec->resize(10, i);
            vectors.push_back(vec);
            std::cout << "  Allocated vector " << i << "\n";
        }
        
        std::cout << "3. Deallocating vectors...\n";
        for (auto* vec : vectors) {
            vector_allocator.destroy(vec);
            vector_allocator.deallocate(vec, 1);
        }
        
        std::cout << "4. Printing statistics...\n";
        vector_allocator.printStatsReport();
        
        std::cout << "5. About to exit scope (destructor will be called)...\n";
    } // vector_allocator 析构函数在这里被调用
    
    std::cout << "6. Successfully exited scope!\n";
}

void testSimpleAllocatorDestruction() {
    std::cout << "\n=== Testing Simple Allocator Destruction ===\n";
    
    {
        std::cout << "1. Creating PoolAllocator<int>...\n";
        PoolAllocator<int> int_allocator;
        
        std::cout << "2. Allocating some integers...\n";
        std::vector<int*> numbers;
        
        for (int i = 0; i < 5; ++i) {
            int* num = int_allocator.allocate(1);
            int_allocator.construct(num, i * 10);
            numbers.push_back(num);
            std::cout << "  Allocated number " << *num << "\n";
        }
        
        std::cout << "3. Deallocating integers...\n";
        for (auto* num : numbers) {
            int_allocator.destroy(num);
            int_allocator.deallocate(num, 1);
        }
        
        std::cout << "4. Printing statistics...\n";
        int_allocator.printStatsReport();
        
        std::cout << "5. About to exit scope (destructor will be called)...\n";
    } // int_allocator 析构函数在这里被调用
    
    std::cout << "6. Successfully exited scope!\n";
}

int main() {
    try {
        // 初始化日志系统
        Logger::getInstance().initialize("minimal_destructor_test.log", Logger::Level::INFO);
        
        std::cout << "=== Minimal Destructor Test ===\n";
        std::cout << "Testing PoolAllocator destruction issues.\n";
        
        // 测试简单类型分配器
        testSimpleAllocatorDestruction();
        
        // 测试复杂类型分配器（这里可能会卡住）
        testVectorAllocatorDestruction();
        
        std::cout << "\n=== All Tests Completed Successfully ===\n";
        
        return 0;
        
    } catch (const std::exception& e) {
        std::cerr << "Test failed with exception: " << e.what() << std::endl;
        return 1;
    } catch (...) {
        std::cerr << "Test failed with unknown exception" << std::endl;
        return 1;
    }
}