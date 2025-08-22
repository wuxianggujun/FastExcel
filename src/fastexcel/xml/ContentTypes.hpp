#include "fastexcel/utils/Logger.hpp"
#pragma once

#include "fastexcel/xml/XMLStreamWriter.hpp"
#include <string>
#include <vector>
#include <functional>

namespace fastexcel {
namespace xml {

class ContentTypes {
public:
    ContentTypes() = default;
    ~ContentTypes() = default;
    
    // 添加默认内容类型
    void addDefault(const std::string& extension, const std::string& content_type);
    
    // 添加覆盖内容类型
    void addOverride(const std::string& part_name, const std::string& content_type);
    
    // 生成XML内容到回调函数（流式写入）
    void generate(const std::function<void(const char*, size_t)>& callback) const;
    
    // 生成XML内容到文件（流式写入）
    void generateToFile(const std::string& filename) const;
    
    // 清空内容
    void clear();
    
    // 添加默认的Excel内容类型
    void addExcelDefaults();
    
private:
    struct DefaultType {
        std::string extension;
        std::string content_type;
    };
    
    struct OverrideType {
        std::string part_name;
        std::string content_type;
    };
    
    std::vector<DefaultType> default_types_;
    std::vector<OverrideType> override_types_;
};

}} // namespace fastexcel::xml
