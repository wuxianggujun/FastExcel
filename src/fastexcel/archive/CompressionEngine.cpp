#include "CompressionEngine.hpp"
#include "ZlibEngine.hpp"

#ifdef FASTEXCEL_HAS_LIBDEFLATE
#include "LibDeflateEngine.hpp"
#endif

#include <stdexcept>
#include <algorithm>
#include <vector>

namespace fastexcel {
namespace archive {

std::unique_ptr<CompressionEngine> CompressionEngine::create(Backend backend, int compression_level) {
    switch (backend) {
        case Backend::ZLIB:
            return std::make_unique<ZlibEngine>(compression_level);
        
        case Backend::LIBDEFLATE:
#ifdef FASTEXCEL_HAS_LIBDEFLATE
            return std::make_unique<LibDeflateEngine>(compression_level);
#else
            throw std::runtime_error("LibDeflate backend not compiled in. Rebuild with FASTEXCEL_USE_LIBDEFLATE=ON");
#endif
        
        default:
            throw std::invalid_argument("Unknown compression backend");
    }
}

std::vector<CompressionEngine::Backend> CompressionEngine::getAvailableBackends() {
    std::vector<Backend> backends;
    
    // zlib 总是可用的
    backends.push_back(Backend::ZLIB);
    
    // 检查 libdeflate 是否可用
#ifdef FASTEXCEL_HAS_LIBDEFLATE
    backends.push_back(Backend::LIBDEFLATE);
#endif
    
    return backends;
}

std::string CompressionEngine::backendToString(Backend backend) {
    switch (backend) {
        case Backend::ZLIB:
            return "zlib";
        case Backend::LIBDEFLATE:
            return "libdeflate";
        default:
            return "unknown";
    }
}

CompressionEngine::Backend CompressionEngine::stringToBackend(const std::string& name) {
    std::string lower_name = name;
    std::transform(lower_name.begin(), lower_name.end(), lower_name.begin(), ::tolower);
    
    if (lower_name == "zlib") {
        return Backend::ZLIB;
    } else if (lower_name == "libdeflate") {
        return Backend::LIBDEFLATE;
    } else {
        throw std::invalid_argument("Unknown backend name: " + name);
    }
}

CompressionEngine::Backend selectOptimalBackend(size_t input_size, int compression_level) {
#ifdef FASTEXCEL_HAS_LIBDEFLATE
    // 智能选择逻辑：libdeflate 在大多数情况下都更快
    
    // 大文件优先使用 libdeflate（性能优势更明显）
    if (input_size > 1024 * 1024) {  // > 1MB
        return CompressionEngine::Backend::LIBDEFLATE;
    }
    
    // 低压缩级别时 libdeflate 优势更明显
    if (compression_level <= 3) {
        return CompressionEngine::Backend::LIBDEFLATE;
    }
    
    // 中等大小文件也倾向于使用 libdeflate
    if (input_size > 64 * 1024) {  // > 64KB
        return CompressionEngine::Backend::LIBDEFLATE;
    }
    
    // 小文件使用 zlib（初始化开销相对较小）
    return CompressionEngine::Backend::ZLIB;
#else
    // 如果没有 libdeflate，只能使用 zlib
    (void)input_size;
    (void)compression_level;
    return CompressionEngine::Backend::ZLIB;
#endif
}

} // namespace archive
} // namespace fastexcel