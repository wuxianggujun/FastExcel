#include "fastexcel/utils/Logger.hpp"
#pragma once

#include "fastexcel/core/Path.hpp"
#include <string>
#include <vector>

namespace fastexcel {
namespace opc {

/**
 * @brief 包管理器接口
 * 遵循接口隔离原则 (ISP)：定义清晰的包操作契约
 */
class IPackageManager {
public:
    virtual ~IPackageManager() = default;
    
    // 读取操作
    virtual bool openForReading(const core::Path& path) = 0;
    virtual std::string readPart(const std::string& part_name) = 0;
    virtual bool partExists(const std::string& part_name) const = 0;
    virtual std::vector<std::string> listParts() const = 0;
    
    // 写入操作
    virtual bool openForWriting(const core::Path& path) = 0;
    virtual bool writePart(const std::string& part_name, const std::string& content) = 0;
    virtual bool removePart(const std::string& part_name) = 0;
    virtual bool commit() = 0;
    
    // 状态查询
    virtual bool isReadable() const = 0;
    virtual bool isWritable() const = 0;
    virtual size_t getPartCount() const = 0;
};

} // namespace opc
} // namespace fastexcel
