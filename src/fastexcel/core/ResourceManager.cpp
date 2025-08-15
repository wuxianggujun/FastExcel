#include "fastexcel/core/ResourceManager.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/IFileWriter.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <filesystem>
#include <fstream>
#include <random>
#include <chrono>
#include <sstream>

namespace fs = std::filesystem;

namespace fastexcel {
namespace core {

// 静态成员定义 - 核心组件必须重建
const std::vector<std::string> ResourceManager::CORE_COMPONENTS_TO_REBUILD = {
    "xl/workbook.xml",
    "xl/worksheets/",
    "xl/sharedStrings.xml",
    "xl/styles.xml",
    "xl/_rels/",
    "[Content_Types].xml",
    "_rels/.rels",
    "docProps/app.xml",
    "docProps/core.xml",
    "xl/calcChain.xml",
    "xl/theme/"
};

// 安全透传的组件
const std::vector<std::string> ResourceManager::SAFE_TO_PASSTHROUGH = {
    "xl/media/",          // 图片等媒体文件
    "xl/drawings/",       // 绘图
    "xl/charts/",         // 图表
    "xl/embeddings/",     // 嵌入对象
    "xl/vbaProject.bin",  // VBA项目
    "xl/ctrlProps/",      // 控件属性
    "xl/customXml/",      // 自定义XML
    "xl/externalLinks/",  // 外部链接
    "xl/pivotCache/",     // 数据透视表缓存
    "xl/pivotTables/",    // 数据透视表
    "xl/queryTables/",    // 查询表
    "xl/slicerCaches/",   // 切片器缓存
    "xl/slicers/",        // 切片器
    "xl/tables/",         // 表格
    "xl/timelines/",      // 时间线
    "xl/model/"           // 数据模型
};

ResourceManager::ResourceManager()
    : mode_(Mode::WRITE_NEW)
    , is_open_(false)
    , delayed_write_mode_(true) {
}

ResourceManager::ResourceManager(const Path& path, Mode mode)
    : filename_(path.string())
    , mode_(mode)
    , is_open_(false)
    , delayed_write_mode_(true) {
}

ResourceManager::~ResourceManager() {
    if (is_open_) {
        close();
    }
}

bool ResourceManager::open(bool create_if_not_exists) {
    if (is_open_) {
        FASTEXCEL_LOG_WARN("ResourceManager is already open");
        return true;
    }
    
    // 检查文件是否存在
    bool file_exists = fs::exists(filename_);
    
    if (!file_exists && !create_if_not_exists) {
        FASTEXCEL_LOG_ERROR("File does not exist: " + filename_);
        return false;
    }
    
    // 根据模式决定是否立即打开文件管理器
    if (mode_ == Mode::READ_ONLY || !delayed_write_mode_) {
        return openInternal(mode_ != Mode::READ_ONLY);
    }
    
    // 延迟写入模式下，仅标记为已打开
    is_open_ = true;
    FASTEXCEL_LOG_DEBUG("ResourceManager opened in delayed mode for: " + filename_);
    return true;
}

bool ResourceManager::openInternal(bool for_writing) {
    if (!file_manager_) {
        file_manager_ = std::make_unique<archive::FileManager>(Path(filename_));
    }
    
    bool success = file_manager_->open(!for_writing); // open(create = !for_writing)
    
    
    if (success) {
        is_open_ = true;
        FASTEXCEL_LOG_DEBUG("FileManager opened: " + filename_);
    } else {
        FASTEXCEL_LOG_ERROR("Failed to open FileManager: " + filename_);
    }
    
    return success;
}

bool ResourceManager::close() {
    if (!is_open_) {
        return true;
    }
    
    bool success = true;
    if (file_manager_ && file_manager_->isOpen()) {
        success = file_manager_->close();
    }
    
    is_open_ = false;
    file_manager_.reset();
    
    FASTEXCEL_LOG_DEBUG("ResourceManager closed: " + filename_);
    return success;
}

bool ResourceManager::prepareForEditing(const Path& path, const std::string& original_path) {
    if (is_open_) {
        FASTEXCEL_LOG_WARN("ResourceManager is already open, closing first");
        close();
    }
    
    filename_ = path.string();
    original_package_path_ = original_path;
    mode_ = Mode::EDIT_EXISTING;
    
    // 延迟打开模式
    is_open_ = true;
    FASTEXCEL_LOG_DEBUG("Prepared for editing: " + filename_ + " (original: " + original_path + ")");
    return true;
}

bool ResourceManager::atomicSave(const Workbook* workbook, const SaveStrategy& strategy) {
    if (!workbook) {
        FASTEXCEL_LOG_ERROR("Workbook is null");
        return false;
    }
    
    // 处理同文件保存的特殊情况
    if (mode_ == Mode::EDIT_EXISTING && 
        fs::equivalent(filename_, original_package_path_)) {
        return handleSameFileSave(strategy);
    }
    
    // 常规保存流程
    return saveInternal(workbook, strategy);
}

bool ResourceManager::saveAs(const Path& new_path, const Workbook* workbook) {
    if (!workbook) {
        FASTEXCEL_LOG_ERROR("Workbook is null");
        return false;
    }
    
    // 临时更改文件名
    std::string old_filename = filename_;
    filename_ = new_path.string();
    
    SaveStrategy strategy;
    strategy.use_temp_file = false;  // 另存为不需要临时文件
    
    bool success = saveInternal(workbook, strategy);
    
    if (!success) {
        filename_ = old_filename;  // 恢复原文件名
    }
    
    return success;
}

bool ResourceManager::saveInternal(const Workbook* workbook, const SaveStrategy& strategy) {
    // 确保文件管理器已打开
    if (!file_manager_ || !file_manager_->isOpen()) {
        if (!openInternal(true)) {
            FASTEXCEL_LOG_ERROR("Failed to open FileManager for writing");
            return false;
        }
    }
    
    try {
        // 1. 如果需要透传复制，先复制非核心组件
        if (needsPassthrough()) {
            auto skip_prefixes = getSkipPrefixes(strategy);
            if (!copyFromOriginalPackage(Path(original_package_path_), skip_prefixes)) {
                FASTEXCEL_LOG_WARN("Failed to copy from original package, continuing anyway");
            }
        }
        
        // 2. 保存工作簿内容（这里需要调用Workbook的保存逻辑）
        // 注意：这里需要Workbook提供一个接口来写入所有组件
        // 暂时使用伪代码表示
        // workbook->writeToResourceManager(this);
        
        // 3. 关闭文件管理器
        if (!file_manager_->close()) {
            FASTEXCEL_LOG_ERROR("Failed to close FileManager");
            return false;
        }
        
        FASTEXCEL_LOG_INFO("Successfully saved: " + filename_);
        return true;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Exception during save: " + std::string(e.what()));
        return false;
    }
}

bool ResourceManager::handleSameFileSave(const SaveStrategy& strategy) {
    if (!strategy.use_temp_file) {
        // 直接覆盖原文件
        return saveInternal(nullptr, strategy);
    }
    
    // 使用临时文件进行原子保存
    std::string temp_path = createTempPath(filename_);
    std::string backup_path = filename_ + ".bak";
    
    try {
        // 1. 保存到临时文件
        std::string original_filename = filename_;
        filename_ = temp_path;
        
        if (!saveInternal(nullptr, strategy)) {
            filename_ = original_filename;
            fs::remove(temp_path);
            return false;
        }
        
        filename_ = original_filename;
        
        // 2. 创建备份（如果需要）
        if (strategy.preserve_backup && fs::exists(filename_)) {
            fs::copy_file(filename_, backup_path, 
                         fs::copy_options::overwrite_existing);
        }
        
        // 3. 原子替换
        if (!atomicReplace(Path(temp_path), Path(filename_))) {
            // 恢复备份
            if (strategy.preserve_backup && fs::exists(backup_path)) {
                fs::rename(backup_path, filename_);
            }
            return false;
        }
        
        // 4. 清理备份（如果不需要保留）
        if (!strategy.preserve_backup && fs::exists(backup_path)) {
            fs::remove(backup_path);
        }
        
        FASTEXCEL_LOG_INFO("Successfully saved with atomic replace: " + filename_);
        return true;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Exception during atomic save: " + std::string(e.what()));
        cleanupTempFiles({temp_path, backup_path});
        return false;
    }
}

bool ResourceManager::copyFromOriginalPackage(const Path& source_path, 
                                             const std::vector<std::string>& skip_prefixes) {
    if (!file_manager_ || !file_manager_->isOpen()) {
        FASTEXCEL_LOG_ERROR("FileManager is not open for writing");
        return false;
    }
    
    try {
        // 打开源文件进行读取
        archive::FileManager source_manager(source_path);
        if (!source_manager.open(false)) { // false = don't create, just open for reading
            FASTEXCEL_LOG_ERROR("Failed to open source package: " + source_path.string());
            return false;
        }
        
        // 获取所有文件列表
        auto files = source_manager.listFiles();
        
        // 复制非核心组件
        for (const auto& file : files) {
            bool should_skip = false;
            
            // 检查是否应该跳过
            for (const auto& prefix : skip_prefixes) {
                if (file.find(prefix) == 0) {
                    should_skip = true;
                    FASTEXCEL_LOG_DEBUG("Skipping: " + file);
                    break;
                }
            }
            
            if (!should_skip) {
                // 读取并写入文件
                std::string content;
                if (source_manager.readFile(file, content)) {
                    file_manager_->writeFile(file, content);
                    FASTEXCEL_LOG_DEBUG("Copied: " + file);
                }
            }
        }
        
        source_manager.close();
        FASTEXCEL_LOG_INFO("Successfully copied non-core components from: " + source_path.string());
        return true;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Exception during passthrough copy: " + std::string(e.what()));
        return false;
    }
}

bool ResourceManager::smartPassthrough(const Path& source_path, 
                                      bool preserve_media, 
                                      bool preserve_vba) {
    std::vector<std::string> skip_list = CORE_COMPONENTS_TO_REBUILD;
    
    // 根据选项决定是否保留某些组件
    if (!preserve_media) {
        skip_list.push_back("xl/media/");
        skip_list.push_back("xl/drawings/");
    }
    
    if (!preserve_vba) {
        skip_list.push_back("xl/vbaProject.bin");
        skip_list.push_back("xl/ctrlProps/");
    }
    
    return copyFromOriginalPackage(source_path, skip_list);
}

bool ResourceManager::writeFile(const std::string& internal_path, 
                               const std::string& content) {
    if (!file_manager_ || !file_manager_->isOpen()) {
        if (!openInternal(true)) {
            FASTEXCEL_LOG_ERROR("Failed to open FileManager for writing");
            return false;
        }
    }
    
    return file_manager_->writeFile(internal_path, content);
}

bool ResourceManager::writeFile(const std::string& internal_path, 
                               const std::vector<uint8_t>& data) {
    if (!file_manager_ || !file_manager_->isOpen()) {
        if (!openInternal(true)) {
            FASTEXCEL_LOG_ERROR("Failed to open FileManager for writing");
            return false;
        }
    }
    
    return file_manager_->writeFile(internal_path, data);
}

bool ResourceManager::writeFiles(const std::vector<std::pair<std::string, std::string>>& files) {
    if (!file_manager_ || !file_manager_->isOpen()) {
        if (!openInternal(true)) {
            FASTEXCEL_LOG_ERROR("Failed to open FileManager for writing");
            return false;
        }
    }
    
    bool all_success = true;
    for (const auto& [path, content] : files) {
        if (!file_manager_->writeFile(path, content)) {
            FASTEXCEL_LOG_ERROR("Failed to write file: " + path);
            all_success = false;
        }
    }
    
    return all_success;
}

std::unique_ptr<IFileWriter> ResourceManager::createFileWriter(bool use_streaming) {
    // 这里需要实现文件写入器的创建逻辑
    // 由于IFileWriter的具体实现不在当前文件中，我们返回nullptr作为占位符
    FASTEXCEL_LOG_DEBUG("Creating file writer (streaming: " + std::to_string(use_streaming) + ")");
    return nullptr; // 需要具体的IFileWriter实现
}

bool ResourceManager::setCompressionLevel(int level) {
    if (level < 0 || level > 9) {
        FASTEXCEL_LOG_ERROR("Invalid compression level: " + std::to_string(level));
        return false;
    }
    
    if (!file_manager_) {
        file_manager_ = std::make_unique<archive::FileManager>(Path(filename_));
    }
    
    file_manager_->setCompressionLevel(level);
    FASTEXCEL_LOG_DEBUG("Set compression level to: " + std::to_string(level));
    return true;
}

std::string ResourceManager::createTempPath(const std::string& base_path) {
    // 生成唯一的临时文件名
    auto now = std::chrono::system_clock::now();
    auto timestamp = std::chrono::duration_cast<std::chrono::milliseconds>(
        now.time_since_epoch()).count();
    
    std::random_device rd;
    std::mt19937 gen(rd());
    std::uniform_int_distribution<> dis(1000, 9999);
    
    std::stringstream ss;
    ss << base_path << ".tmp_" << timestamp << "_" << dis(gen);
    
    return ss.str();
}

bool ResourceManager::atomicReplace(const Path& temp_path, const Path& target_path) {
    try {
        // Windows和Unix的原子替换策略不同
#ifdef _WIN32
        // Windows下使用MoveFileEx进行原子替换
        if (!fs::exists(temp_path.string())) {
            FASTEXCEL_LOG_ERROR("Temp file does not exist: " + temp_path.string());
            return false;
        }
        
        // 先删除目标文件（Windows不支持直接覆盖）
        if (fs::exists(target_path.string())) {
            fs::remove(target_path.string());
        }
        
        fs::rename(temp_path.string(), target_path.string());
#else
        // Unix/Linux下rename是原子操作
        fs::rename(temp_path.string(), target_path.string());
#endif
        
        FASTEXCEL_LOG_DEBUG("Atomic replace successful: " + target_path.string());
        return true;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Atomic replace failed: " + std::string(e.what()));
        return false;
    }
}

std::vector<std::string> ResourceManager::getSkipPrefixes(const SaveStrategy& strategy) const {
    std::vector<std::string> skip_list;
    
    // 添加默认的核心组件
    skip_list.insert(skip_list.end(), 
                    CORE_COMPONENTS_TO_REBUILD.begin(), 
                    CORE_COMPONENTS_TO_REBUILD.end());
    
    // 添加策略中指定的额外跳过组件
    skip_list.insert(skip_list.end(),
                    strategy.skip_components.begin(),
                    strategy.skip_components.end());
    
    return skip_list;
}

void ResourceManager::cleanupTempFiles(const std::vector<std::string>& temp_files) {
    for (const auto& file : temp_files) {
        try {
            if (fs::exists(file)) {
                fs::remove(file);
                FASTEXCEL_LOG_DEBUG("Cleaned up temp file: " + file);
            }
        } catch (const std::exception& e) {
            FASTEXCEL_LOG_WARN("Failed to clean up temp file: " + file + " - " + e.what());
        }
    }
}

}} // namespace fastexcel::core
