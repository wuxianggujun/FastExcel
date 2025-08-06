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

Result<std::unique_ptr<CompressionEngine>> CompressionEngine::create(Backend backend, int compression_level) {
    try {
        switch (backend) {
            case Backend::ZLIB: {
                auto engine = std::make_unique<ZlibEngine>(compression_level);
                // 显式转换为基类指针
                std::unique_ptr<CompressionEngine> base_engine = std::move(engine);
                return makeExpected(std::move(base_engine));
            }
            
            case Backend::LIBDEFLATE: {
#ifdef FASTEXCEL_HAS_LIBDEFLATE
                auto engine = std::make_unique<LibDeflateEngine>(compression_level);
                std::unique_ptr<CompressionEngine> base_engine = std::move(engine);
                return makeExpected(std::move(base_engine));
#else
                return makeError(ErrorCode::NotImplemented, "LibDeflate backend not compiled in. Rebuild with FASTEXCEL_USE_LIBDEFLATE=ON");
#endif
            }
            
            default:
                return makeError(ErrorCode::InvalidArgument, "Unknown compression backend");
        }
    } catch (const std::exception& e) {
        return makeError(ErrorCode::InternalError, "Failed to create compression engine: " + std::string(e.what()));
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

Result<CompressionEngine::Backend> CompressionEngine::stringToBackend(const std::string& name) {
    std::string lower_name = name;
    std::transform(lower_name.begin(), lower_name.end(), lower_name.begin(), ::tolower);
    
    if (lower_name == "zlib") {
        return makeExpected(Backend::ZLIB);
    } else if (lower_name == "libdeflate") {
        return makeExpected(Backend::LIBDEFLATE);
    } else {
        return makeError(ErrorCode::InvalidArgument, "Unknown backend name: " + name);
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