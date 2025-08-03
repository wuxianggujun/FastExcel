#include "fastexcel/archive/CompressionEngine.hpp"
#include <iostream>
#include <vector>
#include <chrono>
#include <random>
#include <iomanip>
#include <string>
#include <cstring>

using namespace fastexcel::archive;

/**
 * @brief 生成测试数据
 */
std::vector<uint8_t> generateTestData(size_t size, double compressibility = 0.5) {
    std::vector<uint8_t> data(size);
    std::random_device rd;
    std::mt19937 gen(rd());
    
    if (compressibility < 0.1) {
        // 高熵数据（难压缩）
        std::uniform_int_distribution<int> dis(0, 255);
        for (auto& byte : data) {
            byte = static_cast<uint8_t>(dis(gen));
        }
    } else if (compressibility > 0.9) {
        // 低熵数据（易压缩）
        std::uniform_int_distribution<int> dis(0, 10);
        for (auto& byte : data) {
            byte = static_cast<uint8_t>(dis(gen));
        }
    } else {
        // 中等压缩性数据（模拟XML）
        std::string pattern = "<row><c r=\"A1\" t=\"inlineStr\"><is><t>Sample Data ";
        std::uniform_int_distribution<int> num_dis(1000, 9999);
        
        size_t pos = 0;
        while (pos < size) {
            std::string content = pattern + std::to_string(num_dis(gen)) + "</t></is></c></row>\n";
            size_t copy_size = std::min(content.size(), size - pos);
            std::memcpy(data.data() + pos, content.data(), copy_size);
            pos += copy_size;
        }
    }
    
    return data;
}

/**
 * @brief 基准测试单个引擎
 */
void benchmarkEngine(CompressionEngine::Backend backend, 
                    const std::vector<std::vector<uint8_t>>& test_datasets,
                    int compression_level) {
    
    std::cout << "\n=== " << CompressionEngine::backendToString(backend) 
              << " (Level " << compression_level << ") ===" << std::endl;
    
    try {
        auto engine = CompressionEngine::create(backend, compression_level);
        
        double total_input_mb = 0.0;
        double total_output_mb = 0.0;
        double total_time_ms = 0.0;
        size_t successful_compressions = 0;
        
        for (size_t i = 0; i < test_datasets.size(); ++i) {
            const auto& input_data = test_datasets[i];
            
            // 分配输出缓冲区
            size_t max_output_size = engine->getMaxCompressedSize(input_data.size());
            std::vector<uint8_t> output_data(max_output_size);
            
            // 执行压缩
            auto start_time = std::chrono::high_resolution_clock::now();
            auto result = engine->compress(input_data.data(), input_data.size(),
                                         output_data.data(), output_data.size());
            auto end_time = std::chrono::high_resolution_clock::now();
            
            if (result.success) {
                auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);
                double time_ms = duration.count() / 1000.0;
                
                total_input_mb += input_data.size() / (1024.0 * 1024.0);
                total_output_mb += result.compressed_size / (1024.0 * 1024.0);
                total_time_ms += time_ms;
                successful_compressions++;
                
                std::cout << "Dataset " << (i+1) << ": "
                          << std::fixed << std::setprecision(1)
                          << (input_data.size() / 1024.0) << " KB -> "
                          << (result.compressed_size / 1024.0) << " KB ("
                          << std::setprecision(1) 
                          << (100.0 * result.compressed_size / input_data.size()) << "%) "
                          << std::setprecision(2)
                          << "in " << time_ms << " ms ("
                          << ((input_data.size() / 1024.0 / 1024.0) / (time_ms / 1000.0)) << " MB/s)"
                          << std::endl;
            } else {
                std::cout << "Dataset " << (i+1) << ": FAILED - " << result.error_message << std::endl;
            }
        }
        
        // 总结统计
        if (successful_compressions > 0) {
            double avg_compression_ratio = total_output_mb / total_input_mb;
            double avg_speed = total_input_mb / (total_time_ms / 1000.0);
            
            std::cout << "\n📊 Summary:" << std::endl;
            std::cout << "  Total processed: " << std::fixed << std::setprecision(2) 
                      << total_input_mb << " MB" << std::endl;
            std::cout << "  Average compression ratio: " << std::setprecision(1)
                      << (avg_compression_ratio * 100) << "%" << std::endl;
            std::cout << "  Average speed: " << std::setprecision(2)
                      << avg_speed << " MB/s" << std::endl;
            std::cout << "  Total time: " << std::setprecision(0)
                      << total_time_ms << " ms" << std::endl;
            
            // 显示引擎统计
            auto stats = engine->getStatistics();
            std::cout << "  Engine stats: " << stats.compression_count << " compressions, "
                      << std::setprecision(2) << stats.getAverageSpeed() << " MB/s avg" << std::endl;
        }
        
    } catch (const std::exception& e) {
        std::cout << "❌ Failed to create " << CompressionEngine::backendToString(backend) 
                  << " engine: " << e.what() << std::endl;
    }
}

/**
 * @brief 主基准测试函数
 */
int main() {
    std::cout << "🚀 FastExcel Compression Engine Benchmark" << std::endl;
    std::cout << "==========================================" << std::endl;
    
    // 生成多种测试数据集
    std::vector<std::vector<uint8_t>> test_datasets;
    
    std::cout << "📝 Generating test datasets..." << std::endl;
    
    // 小文件测试
    test_datasets.push_back(generateTestData(64 * 1024, 0.6));      // 64KB XML-like
    test_datasets.push_back(generateTestData(256 * 1024, 0.7));     // 256KB XML-like
    
    // 中等文件测试
    test_datasets.push_back(generateTestData(1024 * 1024, 0.6));    // 1MB XML-like
    test_datasets.push_back(generateTestData(2048 * 1024, 0.5));    // 2MB mixed
    
    // 大文件测试
    test_datasets.push_back(generateTestData(4096 * 1024, 0.6));    // 4MB XML-like
    test_datasets.push_back(generateTestData(8192 * 1024, 0.4));    // 8MB mixed
    
    // 不同压缩性测试
    test_datasets.push_back(generateTestData(1024 * 1024, 0.9));    // 1MB highly compressible
    test_datasets.push_back(generateTestData(1024 * 1024, 0.1));    // 1MB barely compressible
    
    std::cout << "Generated " << test_datasets.size() << " test datasets" << std::endl;
    
    // 获取可用的压缩后端
    auto available_backends = CompressionEngine::getAvailableBackends();
    std::cout << "\n🔧 Available compression backends:" << std::endl;
    for (auto backend : available_backends) {
        std::cout << "  - " << CompressionEngine::backendToString(backend) << std::endl;
    }
    
    // 测试不同压缩级别
    std::vector<int> compression_levels = {1, 3, 6};
    
    for (int level : compression_levels) {
        std::cout << "\n" << std::string(50, '=') << std::endl;
        std::cout << "Testing Compression Level " << level << std::endl;
        std::cout << std::string(50, '=') << std::endl;
        
        for (auto backend : available_backends) {
            benchmarkEngine(backend, test_datasets, level);
        }
    }
    
    std::cout << "\n✅ Benchmark completed!" << std::endl;
    
    return 0;
}