#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/BatchFileWriter.hpp"
#include "fastexcel/core/CustomPropertyManager.hpp"
#include "fastexcel/core/DefinedNameManager.hpp"
#include "fastexcel/core/managers/WorkbookDocumentManager.hpp"
#include "fastexcel/core/managers/WorkbookSecurityManager.hpp"
#include "fastexcel/core/managers/WorkbookDataManager.hpp"
#include "fastexcel/core/ExcelStructureGenerator.hpp"
#include "fastexcel/core/Exception.hpp"
#include "fastexcel/core/Path.hpp"
#include "fastexcel/core/StreamingFileWriter.hpp"
#include "fastexcel/core/StyleTransferContext.hpp"
#include "fastexcel/reader/XLSXReader.hpp"
#include "fastexcel/theme/Theme.hpp"
#include "fastexcel/theme/ThemeParser.hpp"
#include "fastexcel/utils/TimeUtils.hpp"
#include "fastexcel/utils/XMLUtils.hpp"
#include "fastexcel/xml/Relationships.hpp"
#include "fastexcel/xml/SharedStrings.hpp"
#include "fastexcel/xml/StyleSerializer.hpp"
#include "fastexcel/xml/UnifiedXMLGenerator.hpp"
#include "fastexcel/xml/XMLStreamWriter.hpp"

#include <algorithm>
#include <ctime>
#include <fstream>
#include <functional>
#include <iomanip>
#include <sstream>
#include <fmt/format.h>

namespace fastexcel {
namespace core {

// Workbook 实现

std::unique_ptr<Workbook> Workbook::create(const std::string& filepath) {
    auto workbook = std::make_unique<Workbook>(Path(filepath));
    
    // 创建工作簿时设置正确的状态
    workbook->file_source_ = FileSource::NEW_FILE;
    workbook->transitionToState(WorkbookState::CREATING, "Workbook::create()");
    
    // 对于 create() 创建的工作簿，强制设置为新文件
    // 因为我们要完全重写目标文件，无论它是否已存在
    if (workbook->document_manager_) {
        workbook->document_manager_->getDirtyManager()->setIsNewFile(true);
    }
    
    // 🎯 API修复：自动打开工作簿，返回可直接使用的对象
    if (!workbook->open()) {
        FASTEXCEL_LOG_ERROR("Failed to open workbook after creation: {}", filepath);
        return nullptr;
    }
    
    return workbook;
}

std::unique_ptr<Workbook> Workbook::openReadOnly(const std::string& filepath) {
    try {
        Path path(filepath);
        if (!path.exists()) {
            FASTEXCEL_LOG_ERROR("File not found for read-only access: {}", filepath);
            return nullptr;
        }
        
        // 使用XLSXReader读取现有文件
        reader::XLSXReader reader(path);
        auto result = reader.open();
        if (result != core::ErrorCode::Ok) {
            FASTEXCEL_LOG_ERROR("Failed to open XLSX file for reading: {}, error code: {}", filepath, static_cast<int>(result));
            return nullptr;
        }
        
        // 加载工作簿
        std::unique_ptr<core::Workbook> loaded_workbook;
        result = reader.loadWorkbook(loaded_workbook);
        reader.close();
        
        if (result != core::ErrorCode::Ok) {
            FASTEXCEL_LOG_ERROR("Failed to load workbook content: error code: {}", static_cast<int>(result));
            return nullptr;
        }
        
        if (!loaded_workbook) {
            FASTEXCEL_LOG_ERROR("loadWorkbook returned Ok but workbook is nullptr");
            return nullptr;
        }
        
        // 设置读取模式相关标志
        loaded_workbook->file_source_ = FileSource::EXISTING_FILE;
        loaded_workbook->transitionToState(WorkbookState::READING, "Workbook::openReadOnly");
        loaded_workbook->original_package_path_ = filepath;
        
        // ⚠️ 修复：只读模式不应该调用open()方法，因为这会覆盖原文件
        // 只读模式下不需要打开FileManager进行写入操作
        // 如果后续需要另存为，可以在saveAs时再打开FileManager
        
        return loaded_workbook;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Exception while loading workbook for reading: {}, error: {}", filepath, e.what());
        return nullptr;
    }
}

std::unique_ptr<Workbook> Workbook::openEditable(const std::string& filepath) {
    try {
        Path path(filepath);
        if (!path.exists()) {
            FASTEXCEL_LOG_ERROR("File not found for editing: {}", filepath);
            return nullptr;
        }
        
        // 使用XLSXReader读取现有文件
        reader::XLSXReader reader(path);
        auto result = reader.open();
        if (result != core::ErrorCode::Ok) {
            FASTEXCEL_LOG_ERROR("Failed to open XLSX file for editing: {}, error code: {}", filepath, static_cast<int>(result));
            return nullptr;
        }
        
        // 加载工作簿
        std::unique_ptr<core::Workbook> loaded_workbook;
        result = reader.loadWorkbook(loaded_workbook);
        reader.close();
        
        if (result != core::ErrorCode::Ok) {
            FASTEXCEL_LOG_ERROR("Failed to load workbook from file: {}, error code: {}", filepath, static_cast<int>(result));
            return nullptr;
        }
        
        // 标记来源以便保存时进行未修改部件的保真写回
        if (loaded_workbook) {
            loaded_workbook->transitionToState(WorkbookState::EDITING, "openEditable()");
            loaded_workbook->original_package_path_ = filepath;
            loaded_workbook->file_source_ = FileSource::EXISTING_FILE;
            
            // 为编辑模式准备FileManager（但不立即打开写入，避免覆盖原文件）
            loaded_workbook->filename_ = filepath;
            loaded_workbook->file_manager_ = std::make_unique<archive::FileManager>(path);
            
            // ⚠️ 重要修复：编辑模式不应该立即覆盖原文件！
            // 只有在调用save()或saveAs()时才打开FileManager进行写入
            // if (!loaded_workbook->file_manager_->open(true)) {
            //     FASTEXCEL_LOG_ERROR("Failed to open FileManager for editing: {}", filepath);
            //     return nullptr;
            // }
        }
        
        FASTEXCEL_LOG_INFO("Successfully loaded workbook for editing: {}", filepath);
        return loaded_workbook;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Exception while loading workbook for editing: {}, error: {}", filepath, e.what());
        return nullptr;
    }
}

Workbook::Workbook(const Path& path) : filename_(path.string()) {
    // 使用异常安全的构造模式
    utils::ResourceManager resource_manager;
    
    try {
        // 检查是否为内存模式（任何以::memory::开头的路径）
        if (path.string().find("::memory::") == 0) {
            // 内存模式：不创建FileManager，保持纯内存操作
            file_manager_ = nullptr;
            FASTEXCEL_LOG_DEBUG("Created workbook in memory mode: {}", filename_);
        } else {
            // 文件模式：创建FileManager处理文件操作
            file_manager_ = std::make_unique<archive::FileManager>(path);
            resource_manager.addResource(file_manager_);
        }
        
        format_repo_ = std::make_unique<FormatRepository>();
        resource_manager.addResource(format_repo_);
        
        // 初始化共享字符串表
        shared_string_table_ = std::make_unique<SharedStringTable>();
        resource_manager.addResource(shared_string_table_);
        
        // 初始化专门管理器（职责分离）
        document_manager_ = std::make_unique<WorkbookDocumentManager>(this);
        resource_manager.addResource(document_manager_);
        
        security_manager_ = std::make_unique<WorkbookSecurityManager>(this);
        resource_manager.addResource(security_manager_);
        
        // 初始化工作表管理器（避免空指针访问）
        worksheet_manager_ = std::make_unique<WorksheetManager>(this);

        data_manager_ = std::make_unique<WorkbookDataManager>(this);
        resource_manager.addResource(data_manager_);
        
        // 初始化DocumentManager的DirtyManager
        document_manager_->getDirtyManager()->setIsNewFile(!path.exists()); // 如果文件不存在，则是新文件
        
        // 设置默认文档属性
        document_manager_->setAuthor("FastExcel");
        document_manager_->setCompany("FastExcel Library");
        
        // 已移除：Workbook 层面的内存统计（统一使用内存池统计）
        
        // 共享字符串默认开启（WorkbookOptions 默认 use_shared_strings = true）
        FASTEXCEL_LOG_DEBUG("Workbook options: use_shared_strings={} (default)", options_.use_shared_strings);

        // 成功构造，取消清理
        resource_manager.release();
        
        FASTEXCEL_LOG_DEBUG("Workbook constructed with memory optimizations: {}", filename_);
        
    } catch (...) {
        // resource_manager会自动清理已分配的资源
        FASTEXCEL_LOG_ERROR("Failed to construct Workbook: {}", path.string());
        throw;
    }
}

Workbook::~Workbook() {
    try {
        // 已移除：Workbook 层面的统计打印
        
        close();
        
        // 清理内存池（统一内存管理器会自动处理）
        if (memory_manager_) {
            memory_manager_->clear();
        }
        
    } catch (...) {
        // 析构函数中不抛出异常
        FASTEXCEL_LOG_ERROR("Error during Workbook destruction");
    }
}

// 文件操作

bool Workbook::open() {
    // 内存模式无需文件操作
    if (!file_manager_) {
        FASTEXCEL_LOG_DEBUG("Memory workbook opened: {}", filename_);
        return true;
    }
    
    // 文件模式需要打开FileManager
    bool success = file_manager_->open(true);
    if (success) {
        FASTEXCEL_LOG_INFO("Workbook opened: {}", filename_);
    }
    
    return success;
}

bool Workbook::save() {
    // 运行时检查：只读模式不能保存
    ensureEditable("save");
    
    // 检查关键组件是否存在
    if (!file_manager_) {
        FASTEXCEL_LOG_ERROR("Cannot save: FileManager is null");
        return false;
    }
    
    try {
        // 使用 TimeUtils 更新修改时间（通过DocumentManager）
        if (document_manager_) {
            document_manager_->updateModifiedTime();
        }
        
        // 确保FileManager已打开，如果没有则打开它
        if (file_manager_ && !file_manager_->isOpen()) {
            FASTEXCEL_LOG_DEBUG("FileManager not open, opening for save operation");
            if (!file_manager_->open(true)) {
                FASTEXCEL_LOG_ERROR("Failed to open FileManager for save operation");
                return false;
            }
        }
        
        // 设置ZIP压缩级别 (添加空指针检查)
        if (file_manager_ && file_manager_->isOpen()) {
            if (!file_manager_->setCompressionLevel(options_.compression_level)) {
                FASTEXCEL_LOG_WARN("Failed to set compression level to {}", options_.compression_level);
            } else {
                FASTEXCEL_LOG_DEBUG("Set ZIP compression level to {}", options_.compression_level);
            }
        } else {
            FASTEXCEL_LOG_ERROR("Cannot save: FileManager is not open");
            return false;
        }
        
        // 预先收集所有字符串，避免动态修改
        if (options_.use_shared_strings) {
            FASTEXCEL_LOG_DEBUG("SharedStrings enabled - pre-collecting all strings from worksheets");
            collectSharedStrings();  // 预先收集所有字符串
            FASTEXCEL_LOG_DEBUG("Collected {} unique strings in SharedStringTable", 
                      shared_string_table_ ? shared_string_table_->getStringCount() : 0);
        } else {
            FASTEXCEL_LOG_DEBUG("SharedStrings disabled for performance");
            if (shared_string_table_) shared_string_table_->clear();
        }
        // 标记需要生成共享字符串，避免DirtyManager不知情导致跳过生成
        if (options_.use_shared_strings) {
            if (auto* dm = getDirtyManager()) {
                dm->markSharedStringsDirty();
            }
        }
        
        // 编辑模式下，先将原包中未被我们生成的条目拷贝过来（绘图、图片、打印设置等）
        // 检查是否是保存到同一文件，避免文件锁定问题
        if (isPassThroughEditMode() && !original_package_path_.empty() && file_manager_ && file_manager_->isOpen()) {
            // 检查是否保存到同一文件
            bool is_same_file = (original_package_path_ == filename_);
            
            if (is_same_file) {
                // 保存到同一文件：先关闭当前FileManager，复制原文件到临时位置
                FASTEXCEL_LOG_DEBUG("Saving to same file, creating temporary backup for resource preservation");
                
                std::string temp_backup = fmt::format("{}.tmp_backup_{}", original_package_path_, static_cast<long long>(std::time(nullptr)));
                core::Path source_path(original_package_path_);
                core::Path temp_path(temp_backup);
                
                try {
                    // 关闭当前FileManager以释放文件锁定
                    file_manager_->close();
                    
                    // 复制原文件到临时位置
                    if (temp_path.exists()) {
                        temp_path.remove();
                    }
                    source_path.copyTo(temp_path);
                    
                    // 重新打开FileManager用于写入
                    if (!file_manager_->open(true)) {
                        FASTEXCEL_LOG_ERROR("Failed to reopen FileManager after backup creation");
                        // 清理临时文件
                        if (temp_path.exists()) temp_path.remove();
                        return false;
                    }
                    
                    // 从临时文件复制内容，跳过核心部件（将被重新生成）
                    std::vector<std::string> skip_prefixes = {
                        "[Content_Types].xml",
                        "_rels/",
                        "xl/workbook.xml",
                        "xl/_rels/",
                        "xl/styles.xml",
                        "xl/sharedStrings.xml",
                        "xl/worksheets/",
                        "xl/theme/",
                        "docProps/"  // 文档属性也重新生成
                    };
                    file_manager_->copyFromExistingPackage(temp_path, skip_prefixes);
                    
                    // 清理临时文件
                    if (temp_path.exists()) {
                        temp_path.remove();
                        FASTEXCEL_LOG_DEBUG("Removed temporary backup: {}", temp_backup);
                    }
                    
                } catch (const std::exception& e) {
                    FASTEXCEL_LOG_ERROR("Failed to handle same-file save: {}", e.what());
                    // 清理临时文件
                    if (temp_path.exists()) temp_path.remove();
                    // 尝试重新打开FileManager
                    file_manager_->open(true);
                    // 继续执行，不复制原内容
                }
            } else {
                // 保存到不同文件：正常复制，但跳过核心部件
                std::vector<std::string> skip_prefixes = {
                    "[Content_Types].xml",
                    "_rels/",
                    "xl/workbook.xml",
                    "xl/_rels/",
                    "xl/styles.xml",
                    "xl/sharedStrings.xml",
                    "xl/worksheets/",
                    "xl/theme/",
                    "docProps/"  // 文档属性也重新生成
                };
                file_manager_->copyFromExistingPackage(core::Path(original_package_path_), skip_prefixes);
            }
        }

        // 生成Excel文件结构（会覆盖我们管理的核心部件）
        if (!generateExcelStructure()) {
            FASTEXCEL_LOG_ERROR("Failed to generate Excel structure");
            return false;
        }

        FASTEXCEL_LOG_INFO("Workbook saved successfully: {}", filename_);
        return true;
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Failed to save workbook: {}", e.what());
        return false;
    }
}

bool Workbook::saveAs(const std::string& filename) {
    // 运行时检查：只读模式不能保存
    ensureEditable("saveAs");
    
    std::string old_filename = filename_;
    std::string original_source = original_package_path_;
    bool was_from_existing = (file_source_ == FileSource::EXISTING_FILE);

    // 检查是否保存到同一个文件
    bool is_same_file = (filename == old_filename) || (filename == original_source);
    
    if (is_same_file && was_from_existing && !original_source.empty()) {
        // 如果保存到同一个文件，需要先复制原文件到临时位置
        FASTEXCEL_LOG_INFO("Saving to same file, creating temporary backup for resource preservation");
        
        // 创建临时文件路径
        std::string temp_backup = fmt::format("{}.tmp_backup", original_source);
        core::Path source_path(original_source);
        core::Path temp_path(temp_backup);
        
        // 复制原文件到临时位置
        try {
            if (temp_path.exists()) {
                temp_path.remove();
            }
            source_path.copyTo(temp_path);
            original_package_path_ = temp_backup;  // 更新源路径为临时文件
            FASTEXCEL_LOG_DEBUG("Created temporary backup: {}", temp_backup);
        } catch (const std::exception& e) {
            FASTEXCEL_LOG_ERROR("Failed to create temporary backup: {}", e.what());
            return false;
        }
    }

    filename_ = filename;
    
    // 重新创建文件管理器
    file_manager_ = std::make_unique<archive::FileManager>(core::Path(filename));
    
    if (!file_manager_->open(true)) {
        // 恢复原文件名
        filename_ = old_filename;
        file_manager_ = std::make_unique<archive::FileManager>(core::Path(old_filename));
        
        // 如果创建了临时文件，删除它
        if (is_same_file && original_package_path_.find(".tmp_backup") != std::string::npos) {
            core::Path temp_path(original_package_path_);
            if (temp_path.exists()) {
                temp_path.remove();
            }
            original_package_path_ = original_source;  // 恢复原路径
        }
        return false;
    }

    // 在另存为场景下，如果当前工作簿是从现有包打开的，那么保留 original_package_path_ 用于拷贝未修改部件
    // file_source_ 状态保持不变，已经在之前设置好了
    // original_package_path_ 已经在上面设置好了（可能是临时文件或原始文件）
    
    bool save_result = save();
    
    // 清理临时文件（如果有）
    if (is_same_file && original_package_path_.find(".tmp_backup") != std::string::npos) {
        core::Path temp_path(original_package_path_);
        if (temp_path.exists()) {
            temp_path.remove();
            FASTEXCEL_LOG_DEBUG("Removed temporary backup: {}", original_package_path_);
        }
        original_package_path_ = original_source;  // 恢复原路径
    }
    
    return save_result;
}

bool Workbook::isOpen() const {
    // 检查工作簿是否处于可用状态（EDITING、READING或CREATING模式）
    return state_ == WorkbookState::EDITING || 
           state_ == WorkbookState::READING || 
           state_ == WorkbookState::CREATING;
}

bool Workbook::close() {
    // 幂等性检查：避免重复关闭
    if (state_ == WorkbookState::CLOSED) {
        return true;  // 已经关闭，直接返回成功
    }
    
    // 内存模式只需要重置状态
    if (!file_manager_) {
        FASTEXCEL_LOG_DEBUG("Memory workbook closed: {}", filename_);
    } else {
        // 文件模式需要关闭FileManager
        file_manager_->close();
        FASTEXCEL_LOG_INFO("Workbook closed: {}", filename_);
    }
    
    // 设置为已关闭状态，防止重复关闭
    state_ = WorkbookState::CLOSED;
    return true;
}

// 文档属性（对外便捷接口实现，转发到 WorkbookDocumentManager）
void Workbook::setDocumentProperties(const std::string& title,
                                     const std::string& subject,
                                     const std::string& author,
                                     const std::string& company,
                                     const std::string& comments) {
    if (!document_manager_) return;
    // 新API：设置标题/主题/作者/公司/注释
    document_manager_->setDocumentProperties(title, subject, author, company, comments);
    // 兼容旧示例：将第5个参数同时作为“类别”写入
    if (!comments.empty()) {
        document_manager_->setCategory(comments);
    }
}

void Workbook::setApplication(const std::string& application) {
    if (document_manager_) {
        document_manager_->setApplication(application);
    }
}

// 工作表管理

std::shared_ptr<Worksheet> Workbook::addSheet(const std::string& name) {
    // 运行时检查：只读模式不能添加工作表
    ensureEditable("addSheet");
    
    std::string sheet_name;
    if (name.empty()) {
        sheet_name = generateUniqueSheetName("Sheet1");
    } else {
        // 检查名称是否已存在，如果存在则生成唯一名称
        if (getSheet(name) != nullptr) {
            sheet_name = generateUniqueSheetName(name);
        } else {
            sheet_name = name;
        }
    }
    
    if (!validateSheetName(sheet_name)) {
        FASTEXCEL_LOG_ERROR("Invalid sheet name: {}", sheet_name);
        return nullptr;
    }
    
    auto worksheet = std::make_shared<Worksheet>(sheet_name, std::shared_ptr<Workbook>(this, [](Workbook*){}), next_sheet_id_++);
    
    // 设置 FormatRepository，启用列宽管理功能
    if (format_repo_) {
        worksheet->setFormatRepository(format_repo_.get());
    }
    
    worksheets_.push_back(worksheet);
    
    // 如果这是第一个工作表，自动设置为激活状态
    if (worksheets_.size() == 1) {
        worksheet->setTabSelected(true);
        active_worksheet_index_ = 0;
        FASTEXCEL_LOG_DEBUG("Added worksheet: {} (activated as first sheet)", sheet_name);
    } else {
        FASTEXCEL_LOG_DEBUG("Added worksheet: {}", sheet_name);
    }
    
    return worksheet;
}

std::shared_ptr<Worksheet> Workbook::insertSheet(size_t index, const std::string& name) {
    // 运行时检查：只读模式不能插入工作表
    ensureEditable("insertSheet");
    
    if (index > worksheets_.size()) {
        index = worksheets_.size();
    }
    
    std::string sheet_name = name.empty() ? generateUniqueSheetName("Sheet1") : name;
    
    if (!validateSheetName(sheet_name)) {
        FASTEXCEL_LOG_ERROR("Invalid sheet name: {}", sheet_name);
        return nullptr;
    }
    
    auto worksheet = std::make_shared<Worksheet>(sheet_name, std::shared_ptr<Workbook>(this, [](Workbook*){}), next_sheet_id_++);
    
    // 设置 FormatRepository，启用列宽管理功能
    if (format_repo_) {
        worksheet->setFormatRepository(format_repo_.get());
    }
    
    worksheets_.insert(worksheets_.begin() + index, worksheet);
    
    FASTEXCEL_LOG_DEBUG("Inserted worksheet: {} at index {}", sheet_name, index);
    return worksheet;
}

bool Workbook::removeSheet(const std::string& name) {
    // 运行时检查：只读模式不能删除工作表
    ensureEditable("removeSheet");
    
    auto it = std::find_if(worksheets_.begin(), worksheets_.end(),
                          [&name](const std::shared_ptr<Worksheet>& ws) {
                              return ws->getName() == name;
                          });
    
    if (it != worksheets_.end()) {
        worksheets_.erase(it);
        FASTEXCEL_LOG_DEBUG("Removed worksheet: {}", name);
        return true;
    }
    
    return false;
}

bool Workbook::removeSheet(size_t index) {
    // 运行时检查：只读模式不能删除工作表
    ensureEditable("removeSheet");
    
    if (index < worksheets_.size()) {
        std::string name = worksheets_[index]->getName();
        worksheets_.erase(worksheets_.begin() + index);
        
        // 更新活动工作表索引
        if (active_worksheet_index_ == index) {
            // 如果删除的是当前活动工作表
            if (worksheets_.empty()) {
                active_worksheet_index_ = 0;  // 没有工作表了
            } else if (active_worksheet_index_ >= worksheets_.size()) {
                active_worksheet_index_ = worksheets_.size() - 1;  // 设置为最后一个
                worksheets_[active_worksheet_index_]->setTabSelected(true);
            } else {
                // 保持当前索引，激活新的工作表
                worksheets_[active_worksheet_index_]->setTabSelected(true);
            }
        } else if (active_worksheet_index_ > index) {
            // 如果删除的工作表在活动工作表之前，索引需要减1
            active_worksheet_index_--;
        }
        
        FASTEXCEL_LOG_DEBUG("Removed worksheet: {} at index {}", name, index);
        return true;
    }
    
    return false;
}

std::shared_ptr<Worksheet> Workbook::getSheet(const std::string& name) {
    auto it = std::find_if(worksheets_.begin(), worksheets_.end(), 
                          [&name](const std::shared_ptr<Worksheet>& ws) {
                              return ws->getName() == name;
                          });
    
    if (it != worksheets_.end()) {
        return *it;
    }
    
    return nullptr;
}

std::shared_ptr<Worksheet> Workbook::getSheet(size_t index) {
    if (index < worksheets_.size()) {
        return worksheets_[index];
    }
    return nullptr;
}

std::shared_ptr<const Worksheet> Workbook::getSheet(const std::string& name) const {
    auto it = std::find_if(worksheets_.begin(), worksheets_.end(),
                          [&name](const std::shared_ptr<Worksheet>& ws) {
                              return ws->getName() == name;
                          });
    
    if (it != worksheets_.end()) {
        return *it;
    }
    
    return nullptr;
}

std::shared_ptr<const Worksheet> Workbook::getSheet(size_t index) const {
    if (index < worksheets_.size()) {
        return worksheets_[index];
    }
    return nullptr;
}

std::vector<std::string> Workbook::getSheetNames() const {
    std::vector<std::string> names;
    names.reserve(worksheets_.size());
    
    for (const auto& worksheet : worksheets_) {
        names.push_back(worksheet->getName());
    }
    
    return names;
}

// 便捷的工作表查找方法
bool Workbook::hasSheet(const std::string& name) const {
    auto it = std::find_if(worksheets_.begin(), worksheets_.end(),
                          [&name](const std::shared_ptr<Worksheet>& ws) {
                              return ws->getName() == name;
                          });
    return it != worksheets_.end();
}

std::shared_ptr<Worksheet> Workbook::findSheet(const std::string& name) {
    auto it = std::find_if(worksheets_.begin(), worksheets_.end(),
                          [&name](const std::shared_ptr<Worksheet>& ws) {
                              return ws->getName() == name;
                          });
    
    if (it != worksheets_.end()) {
        return *it;
    }
    
    return nullptr;
}

std::shared_ptr<const Worksheet> Workbook::findSheet(const std::string& name) const {
    auto it = std::find_if(worksheets_.begin(), worksheets_.end(),
                          [&name](const std::shared_ptr<Worksheet>& ws) {
                              return ws->getName() == name;
                          });
    
    if (it != worksheets_.end()) {
        return *it;
    }
    
    return nullptr;
}

std::vector<std::shared_ptr<Worksheet>> Workbook::getAllSheets() {
    std::vector<std::shared_ptr<Worksheet>> sheets;
    sheets.reserve(worksheets_.size());
    
    for (auto& worksheet : worksheets_) {
        sheets.push_back(worksheet);
    }
    
    return sheets;
}

std::vector<std::shared_ptr<const Worksheet>> Workbook::getAllSheets() const {
    std::vector<std::shared_ptr<const Worksheet>> sheets;
    sheets.reserve(worksheets_.size());
    
    for (const auto& worksheet : worksheets_) {
        sheets.push_back(worksheet);
    }
    
    return sheets;
}

int Workbook::clearAllSheets() {
    ensureEditable("clearAllSheets");
    
    int count = static_cast<int>(worksheets_.size());
    worksheets_.clear();
    
    // 重置工作表ID计数器
    next_sheet_id_ = 1;
    
    // 重置活动工作表索引
    active_worksheet_index_ = 0;
    
    FASTEXCEL_LOG_DEBUG("Cleared all worksheets, removed {} sheets", count);
    
    return count;
}

std::shared_ptr<Worksheet> Workbook::getFirstSheet() {
    if (!worksheets_.empty()) {
        return worksheets_.front();
    }
    return nullptr;
}

std::shared_ptr<const Worksheet> Workbook::getFirstSheet() const {
    if (!worksheets_.empty()) {
        return worksheets_.front();
    }
    return nullptr;
}

std::shared_ptr<Worksheet> Workbook::getLastSheet() {
    if (!worksheets_.empty()) {
        return worksheets_.back();
    }
    return nullptr;
}

std::shared_ptr<const Worksheet> Workbook::getLastSheet() const {
    if (!worksheets_.empty()) {
        return worksheets_.back();
    }
    return nullptr;
}

bool Workbook::renameSheet(const std::string& old_name, const std::string& new_name) {
    auto worksheet = getSheet(old_name);
    if (!worksheet) {
        return false;
    }
    
    if (!validateSheetName(new_name)) {
        return false;
    }
    
    worksheet->setName(new_name);
    FASTEXCEL_LOG_DEBUG("Renamed worksheet: {} -> {}", old_name, new_name);
    return true;
}

bool Workbook::moveSheet(size_t from_index, size_t to_index) {
    if (from_index >= worksheets_.size() || to_index >= worksheets_.size()) {
        return false;
    }
    
    if (from_index == to_index) {
        return true;
    }
    
    auto worksheet = worksheets_[from_index];
    worksheets_.erase(worksheets_.begin() + from_index);
    
    if (to_index > from_index) {
        to_index--;
    }
    
    worksheets_.insert(worksheets_.begin() + to_index, worksheet);
    
    FASTEXCEL_LOG_DEBUG("Moved worksheet from index {} to {}", from_index, to_index);
    return true;
}

std::shared_ptr<Worksheet> Workbook::copyWorksheet(const std::string& source_name, const std::string& new_name) {
    auto source_worksheet = getSheet(source_name);
    if (!source_worksheet) {
        return nullptr;
    }
    
    if (!validateSheetName(new_name)) {
        return nullptr;
    }
    
    // 创建新工作表
    auto new_worksheet = std::make_shared<Worksheet>(new_name, std::shared_ptr<Workbook>(this, [](Workbook*){}), next_sheet_id_++);
    
    // 实现深拷贝逻辑：复制所有单元格、格式、设置等
    try {
        auto [max_row, max_col] = source_worksheet->getUsedRange();
        
        // 复制所有单元格和格式
        for (int row = 0; row <= max_row; ++row) {
            for (int col = 0; col <= max_col; ++col) {
                if (source_worksheet->hasCellAt(row, col)) {
                    // 使用copyCell方法复制单元格及其格式
                    source_worksheet->copyCell(row, col, row, col, true, false);
                    // 将复制的内容设置到新工作表
                    const auto& source_cell = source_worksheet->getCell(row, col);
                    auto& new_cell = new_worksheet->getCell(row, col);
                    
                    // 复制单元格值
                    new_cell = source_cell;
                    
                    // 复制格式
                    auto format = source_cell.getFormatDescriptor();
                    if (format) {
                        new_worksheet->setCellFormat(row, col, format);
                    }
                }
            }
        }
        
        // 复制工作表设置
        new_worksheet->setTabSelected(source_worksheet->isTabSelected());
        
        // 复制列宽和行高
        for (int col = 0; col <= max_col; ++col) {
            auto width_opt = source_worksheet->tryGetColumnWidth(col);
            if (width_opt && *width_opt > 0) {
                new_worksheet->setColumnWidth(col, *width_opt);
            }
        }
        
        for (int row = 0; row <= max_row; ++row) {
            auto height_opt = source_worksheet->tryGetRowHeight(row);
            if (height_opt && *height_opt > 0) {
                new_worksheet->setRowHeight(row, *height_opt);
            }
        }
        
        FASTEXCEL_LOG_DEBUG("Deep copied worksheet content: {} cells from {} to {}", 
                           (max_row + 1) * (max_col + 1), source_name, new_name);
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Failed to copy worksheet content: {}", e.what());
    }
    
    worksheets_.push_back(new_worksheet);
    
    FASTEXCEL_LOG_DEBUG("Copied worksheet: {} -> {}", source_name, new_name);
    return new_worksheet;
}

void Workbook::setActiveWorksheet(size_t index) {
    // 取消所有工作表的选中状态
    for (auto& worksheet : worksheets_) {
        worksheet->setTabSelected(false);
    }
    
    // 设置指定工作表为活动状态
    if (index < worksheets_.size()) {
        worksheets_[index]->setTabSelected(true);
        active_worksheet_index_ = index;  // 更新活动工作表索引
    }
}

std::shared_ptr<Worksheet> Workbook::getActiveWorksheet() {
    if (worksheets_.empty()) {
        return nullptr;
    }
    
    // 确保活动工作表索引在有效范围内
    if (active_worksheet_index_ >= worksheets_.size()) {
        active_worksheet_index_ = 0;
    }
    
    return worksheets_[active_worksheet_index_];
}

std::shared_ptr<const Worksheet> Workbook::getActiveWorksheet() const {
    if (worksheets_.empty()) {
        return nullptr;
    }
    
    // 确保活动工作表索引在有效范围内
    size_t safe_index = (active_worksheet_index_ < worksheets_.size()) ? 
                        active_worksheet_index_ : 0;
    
    return worksheets_[safe_index];
}

// 样式管理

int Workbook::addStyle(const FormatDescriptor& style) {
    return format_repo_->addFormat(style);
}

int Workbook::addStyle(const StyleBuilder& builder) {
    auto format = builder.build();
    return format_repo_->addFormat(format);
}

std::shared_ptr<const FormatDescriptor> Workbook::getStyle(int style_id) const {
    // 从格式仓储中根据ID获取格式描述符
    return format_repo_->getFormat(style_id);
}

int Workbook::getDefaultStyleId() const {
    return format_repo_->getDefaultFormatId();
}

bool Workbook::isValidStyleId(int style_id) const {
    return format_repo_->isValidFormatId(style_id);
}

const FormatRepository& Workbook::getStyles() const {
    return *format_repo_;
}

void Workbook::setThemeXML(const std::string& theme_xml) {
    theme_xml_ = theme_xml;
    theme_dirty_ = true; // 外部显式设置主题XML视为编辑
    FASTEXCEL_LOG_DEBUG("设置自定义主题XML ({} 字节)", theme_xml_.size());
    // 尝试解析为结构化主题对象
    if (!theme_xml_.empty()) {
        auto parsed = theme::ThemeParser::parseFromXML(theme_xml_);
        if (parsed) {
            theme_ = std::move(parsed);
            FASTEXCEL_LOG_DEBUG("主题XML已解析为对象: {}", theme_->getName());
        } else {
            FASTEXCEL_LOG_WARN("主题XML解析失败，保留原始XML");
        }
    }
}

const std::string& Workbook::getThemeXML() const {
    return theme_xml_;
}

void Workbook::setOriginalThemeXML(const std::string& theme_xml) {
    theme_xml_original_ = theme_xml;
    FASTEXCEL_LOG_DEBUG("保存原始主题XML ({} 字节)", theme_xml_original_.size());
    // 同步解析一次，便于后续编辑
    if (!theme_xml_original_.empty()) {
        auto parsed = theme::ThemeParser::parseFromXML(theme_xml_original_);
        if (parsed) {
            theme_ = std::move(parsed);
            FASTEXCEL_LOG_DEBUG("原始主题XML已解析为对象: {}", theme_->getName());
        }
    }
}

void Workbook::setTheme(const theme::Theme& theme) {
    theme_ = std::make_unique<theme::Theme>(theme);
    // 同步XML缓存
    theme_xml_ = theme_->toXML();
    theme_dirty_ = true;
}

void Workbook::setThemeName(const std::string& name) {
    if (!theme_) theme_ = std::make_unique<theme::Theme>();
    theme_->setName(name);
    theme_xml_.clear(); // 让生成时重新序列化
    theme_dirty_ = true;
}

void Workbook::setThemeColor(theme::ThemeColorScheme::ColorType type, const core::Color& color) {
    if (!theme_) theme_ = std::make_unique<theme::Theme>();
    theme_->colors().setColor(type, color);
    theme_xml_.clear();
    theme_dirty_ = true;
}

bool Workbook::setThemeColorByName(const std::string& name, const core::Color& color) {
    if (!theme_) theme_ = std::make_unique<theme::Theme>();
    bool ok = theme_->colors().setColorByName(name, color);
    if (ok) { theme_xml_.clear(); theme_dirty_ = true; }
    return ok;
}

void Workbook::setThemeMajorFontLatin(const std::string& name) {
    if (!theme_) theme_ = std::make_unique<theme::Theme>();
    theme_->fonts().setMajorFontLatin(name);
    theme_xml_.clear();
    theme_dirty_ = true;
}

void Workbook::setThemeMajorFontEastAsia(const std::string& name) {
    if (!theme_) theme_ = std::make_unique<theme::Theme>();
    theme_->fonts().setMajorFontEastAsia(name);
    theme_xml_.clear();
    theme_dirty_ = true;
}

void Workbook::setThemeMajorFontComplex(const std::string& name) {
    if (!theme_) theme_ = std::make_unique<theme::Theme>();
    theme_->fonts().setMajorFontComplex(name);
    theme_xml_.clear();
    theme_dirty_ = true;
}

void Workbook::setThemeMinorFontLatin(const std::string& name) {
    if (!theme_) theme_ = std::make_unique<theme::Theme>();
    theme_->fonts().setMinorFontLatin(name);
    theme_xml_.clear();
    theme_dirty_ = true;
}

void Workbook::setThemeMinorFontEastAsia(const std::string& name) {
    if (!theme_) theme_ = std::make_unique<theme::Theme>();
    theme_->fonts().setMinorFontEastAsia(name);
    theme_xml_.clear();
    theme_dirty_ = true;
}

void Workbook::setThemeMinorFontComplex(const std::string& name) {
    if (!theme_) theme_ = std::make_unique<theme::Theme>();
    theme_->fonts().setMinorFontComplex(name);
    theme_xml_.clear();
    theme_dirty_ = true;
}

// 工作簿选项

void Workbook::setCalcOptions(bool calc_on_load, bool full_calc_on_load) {
    options_.calc_on_load = calc_on_load;
    options_.full_calc_on_load = full_calc_on_load;
}

// 生成控制判定（使用 DirtyManager 进行管理）

bool Workbook::shouldGenerateContentTypes() const {
    auto* dirty_manager = getDirtyManager();
    if (!dirty_manager) return true;
    return dirty_manager->shouldUpdate("[Content_Types].xml");
}

bool Workbook::shouldGenerateRootRels() const {
    auto* dirty_manager = getDirtyManager();
    if (!dirty_manager) return true;
    return dirty_manager->shouldUpdate("_rels/.rels");
}

bool Workbook::shouldGenerateWorkbookCore() const {
    auto* dirty_manager = getDirtyManager();
    if (!dirty_manager) return true;
    return dirty_manager->shouldUpdate("xl/workbook.xml");
}

bool Workbook::shouldGenerateStyles() const {
    // 始终生成样式文件，保证包内引用一致性：
    // - workbook.xml 和 [Content_Types].xml 总是包含对 xl/styles.xml 的引用
    // - 如不生成，将导致包缺少被引用的部件，Excel 打开会提示修复
    // 样式文件很小，生成最小可用样式的成本可以忽略
    return true;
}

bool Workbook::shouldGenerateTheme() const {
    // 只有在确有主题内容时才生成主题文件
    // 避免请求生成主题但ThemeGenerator找不到内容的问题
    if (!theme_xml_.empty() || !theme_xml_original_.empty() || theme_) {
        return true;
    }
    return false; // 没有主题内容，不生成主题文件
}

bool Workbook::shouldGenerateSharedStrings() const {
    FASTEXCEL_LOG_DEBUG("shouldGenerateSharedStrings() called - analyzing conditions");
    
    if (!options_.use_shared_strings) {
        FASTEXCEL_LOG_DEBUG("SharedStrings generation disabled by options_.use_shared_strings = false");
        return false; // 未启用SST
    }
    FASTEXCEL_LOG_DEBUG("options_.use_shared_strings = true, SharedStrings enabled");
    
    auto* dirty_manager = getDirtyManager();
    if (!dirty_manager) {
        FASTEXCEL_LOG_DEBUG("No dirty manager, SharedStrings generation enabled (default true)");
        return true;
    }
    FASTEXCEL_LOG_DEBUG("DirtyManager exists, checking shouldUpdate for xl/sharedStrings.xml");
    
    bool should_update = dirty_manager->shouldUpdate("xl/sharedStrings.xml");
    FASTEXCEL_LOG_DEBUG("DirtyManager shouldUpdate for SharedStrings: {}", should_update);
    
    // 如果 SharedStringTable 有内容但 DirtyManager 认为不需要更新，则强制生成
    if (shared_string_table_) {
        size_t string_count = shared_string_table_->getStringCount();
        FASTEXCEL_LOG_DEBUG("SharedStringTable contains {} strings", string_count);
        
        if (string_count > 0 && !should_update) {
            FASTEXCEL_LOG_DEBUG("FORCE GENERATION: SharedStringTable has {} strings but DirtyManager says no update needed", string_count);
            FASTEXCEL_LOG_DEBUG("This happens when target file exists but we're creating new content with strings");
            FASTEXCEL_LOG_DEBUG("Forcing SharedStrings generation to avoid missing sharedStrings.xml");
            return true; // 强制生成
        }
    } else {
        FASTEXCEL_LOG_DEBUG("SharedStringTable is null");
    }
    
    return should_update;
}

bool Workbook::shouldGenerateDocPropsCore() const {
    auto* dirty_manager = getDirtyManager();
    if (!dirty_manager) return true;
    return dirty_manager->shouldUpdate("docProps/core.xml");
}

bool Workbook::shouldGenerateDocPropsApp() const {
    auto* dirty_manager = getDirtyManager();
    if (!dirty_manager) return true;
    return dirty_manager->shouldUpdate("docProps/app.xml");
}

bool Workbook::shouldGenerateDocPropsCustom() const {
    // 只有在真正有自定义属性时才需要生成 custom.xml
    // 通过 WorkbookDocumentManager 查询自定义属性集合，避免误判
    if (!document_manager_) return false;
    auto props = document_manager_->getAllCustomProperties();
    return !props.empty();
}

bool Workbook::shouldGenerateSheet(size_t index) const {
    auto* dirty_manager = getDirtyManager();
    if (!dirty_manager) return true;
    std::string sheetPart = fmt::format("xl/worksheets/sheet{}.xml", index + 1);
    return dirty_manager->shouldUpdate(sheetPart);
}

bool Workbook::shouldGenerateSheetRels(size_t index) const {
    auto* dirty_manager = getDirtyManager();
    if (!dirty_manager) return true;
    std::string sheetRelsPart = fmt::format("xl/worksheets/_rels/sheet{}.xml.rels", index + 1);
    return dirty_manager->shouldUpdate(sheetRelsPart);
}

// 共享字符串管理

int Workbook::addSharedString(const std::string& str) {
    if (!shared_string_table_) shared_string_table_ = std::make_unique<SharedStringTable>();
    return static_cast<int>(shared_string_table_->addString(str));
}

int Workbook::addSharedStringWithIndex(const std::string& str, int original_index) {
    if (!shared_string_table_) shared_string_table_ = std::make_unique<SharedStringTable>();
    return static_cast<int>(shared_string_table_->addStringWithId(str, original_index));
}

int Workbook::getSharedStringIndex(const std::string& str) const {
    if (!shared_string_table_) return -1;
    return static_cast<int>(shared_string_table_->getStringId(str));
}

const SharedStringTable* Workbook::getSharedStrings() const {
    return shared_string_table_.get();
}

// 内部方法

bool Workbook::generateExcelStructure() {
    // 智能选择生成模式：根据数据量和内存使用情况自动决定
    size_t estimated_memory = estimateMemoryUsage();
    size_t total_cells = getTotalCellCount();
    
    bool use_streaming = false;
    
    // 新的决策逻辑：基于WorkbookMode
    switch (options_.mode) {
        case WorkbookMode::AUTO:
            // 自动模式：根据数据量智能选择
            if (total_cells > options_.auto_mode_cell_threshold ||
                estimated_memory > options_.auto_mode_memory_threshold) {
                use_streaming = true;
                FASTEXCEL_LOG_INFO("Auto-selected streaming mode: {} cells, {}MB estimated memory (thresholds: {} cells, {}MB)",
                        total_cells, estimated_memory / (1024*1024),
                        options_.auto_mode_cell_threshold, options_.auto_mode_memory_threshold / (1024*1024));
            } else {
                use_streaming = false;
                FASTEXCEL_LOG_INFO("Auto-selected batch mode: {} cells, {}MB estimated memory (thresholds: {} cells, {}MB)",
                        total_cells, estimated_memory / (1024*1024),
                        options_.auto_mode_cell_threshold, options_.auto_mode_memory_threshold / (1024*1024));
            }
            break;
            
        case WorkbookMode::BATCH:
            // 强制批量模式
            use_streaming = false;
            FASTEXCEL_LOG_INFO("Using forced batch mode: {} cells, {}MB estimated memory",
                    total_cells, estimated_memory / (1024*1024));
            break;
            
        case WorkbookMode::STREAMING:
            // 强制流式模式
            use_streaming = true;
            FASTEXCEL_LOG_INFO("Using forced streaming mode: {} cells, {}MB estimated memory",
                    total_cells, estimated_memory / (1024*1024));
            break;
    }
    
    // 如果设置了constant_memory，强制使用流式模式
    if (options_.constant_memory) {
        use_streaming = true;
        FASTEXCEL_LOG_INFO("Constant memory mode enabled, forcing streaming mode");
    }
    
    // 实际生成写入
    return generateWithGenerator(use_streaming);
}




// 主题写出逻辑已迁移至 XML 层（ThemeGenerator），此处不再直接输出

// 格式管理内部方法


// 辅助函数

std::string Workbook::generateUniqueSheetName(const std::string& base_name) const {
    // 如果base_name不存在，直接返回
    if (getSheet(base_name) == nullptr) {
        return base_name;
    }
    
    // 如果base_name是"Sheet1"，从"Sheet2"开始尝试
    if (base_name == "Sheet1") {
        int counter = 2;
        std::string name = fmt::format("Sheet{}", counter);
        while (getSheet(name) != nullptr) {
            name = fmt::format("Sheet{}", ++counter);
        }
        return name;
    }
    
    // 对于其他base_name，添加数字后缀
    int suffix_counter = 1;
    std::string name = fmt::format("{}{}", base_name, suffix_counter);
    while (getSheet(name) != nullptr) {
        name = fmt::format("{}{}", base_name, ++suffix_counter);
    }
    
    return name;
}

bool Workbook::validateSheetName(const std::string& name) const {
    // 检查长度
    if (name.empty() || name.length() > 31) {
        return false;
    }
    
    // 检查非法字符
    const std::string invalid_chars = ":\\/?*[]";
    if (name.find_first_of(invalid_chars) != std::string::npos) {
        return false;
    }
    
    // 检查是否以单引号开头或结尾
    if (name.front() == '\'' || name.back() == '\'') {
        return false;
    }
    
    // 不检查是否已存在，因为这个方法也被用于验证新名称
    // 重复名称的检查应该在调用方处理
    
    return true;
}

void Workbook::collectSharedStrings() {
    if (!shared_string_table_) {
        shared_string_table_ = std::make_unique<SharedStringTable>();
    } else {
        shared_string_table_->clear();
    }
    
    for (const auto& worksheet : worksheets_) {
        // 这里需要访问工作表的单元格来收集字符串
        // 简化版本，实际实现需要遍历所有字符串单元格
        auto [max_row, max_col] = worksheet->getUsedRange();
        
        for (int row = 0; row <= max_row; ++row) {
            for (int col = 0; col <= max_col; ++col) {
                if (worksheet->hasCellAt(row, col)) {
                    const auto& cell = worksheet->getCell(row, col);
                    if (cell.isString()) {
                        addSharedString(cell.getStringValue());
                    }
                }
            }
        }
    }
}


std::string Workbook::getWorksheetPath(int sheet_id) const {
    return fmt::format("xl/worksheets/sheet{}.xml", sheet_id);
}

std::string Workbook::getWorksheetRelPath(int sheet_id) const {
    return fmt::format("worksheets/sheet{}.xml", sheet_id);
}


void Workbook::setHighPerformanceMode(bool enable) {
    if (enable) {
        FASTEXCEL_LOG_INFO("Enabling ultra high performance mode (beyond defaults)");
        
        // 进一步优化：使用无压缩模式排除压缩算法影响
        options_.compression_level = 0;  // 无压缩
        
        // 更大的缓冲区
        options_.row_buffer_size = 10000;
        options_.xml_buffer_size = 8 * 1024 * 1024;  // 8MB
        
        // 使用AUTO模式，让系统根据数据量自动选择
        options_.mode = WorkbookMode::AUTO;
        options_.use_shared_strings = true;
        
        // 调整自动模式阈值，更倾向于使用批量模式以获得更好的性能
        options_.auto_mode_cell_threshold = 2000000;  // 200万单元格
        options_.auto_mode_memory_threshold = 200 * 1024 * 1024;  // 200MB
        
        FASTEXCEL_LOG_INFO("Ultra high performance mode configured: Mode=AUTO, Compression=OFF, RowBuffer={}, XMLBuffer={}MB",
                options_.row_buffer_size, options_.xml_buffer_size / (1024*1024));
    } else {
        FASTEXCEL_LOG_INFO("Using standard high performance mode (default settings)");
        
        // 恢复到默认的高性能设置
        options_.mode = WorkbookMode::AUTO;           // 默认自动模式
        options_.use_shared_strings = true;           // 默认启用以匹配Excel格式
        options_.row_buffer_size = 5000;              // 默认较大缓冲
        options_.compression_level = 6;               // 恢复到默认的中等压缩
        options_.xml_buffer_size = 4 * 1024 * 1024;  // 默认4MB
        
        // 恢复默认阈值
        options_.auto_mode_cell_threshold = 1000000;     // 100万单元格
        options_.auto_mode_memory_threshold = 100 * 1024 * 1024; // 100MB
    }
}




// 工作簿编辑功能实现

std::unique_ptr<Workbook> Workbook::open(const Path& path) {
    try {
        // 使用Path的内置文件检查
        if (!path.exists()) {
            FASTEXCEL_LOG_ERROR("File not found for editing: {}", path.string());
            return nullptr;
        }
        
        // 使用XLSXReader读取现有文件
        reader::XLSXReader reader(path);
        auto result = reader.open();
        if (result != core::ErrorCode::Ok) {
            FASTEXCEL_LOG_ERROR("Failed to open XLSX file for reading: {}, error code: {}", path.string(), static_cast<int>(result));
            return nullptr;
        }
        
        // 加载工作簿
        std::unique_ptr<core::Workbook> loaded_workbook;
        result = reader.loadWorkbook(loaded_workbook);
        reader.close();
        
        if (result != core::ErrorCode::Ok) {
            FASTEXCEL_LOG_ERROR("Failed to load workbook from file: {}, error code: {}", path.string(), static_cast<int>(result));
            return nullptr;
        }
        
        // 标记来源以便保存时进行未修改部件的保真写回
        if (loaded_workbook) {
            loaded_workbook->transitionToState(WorkbookState::EDITING, "openForEditing()");
            loaded_workbook->original_package_path_ = path.string();
            // 设置为编辑模式（保持向后兼容）
            loaded_workbook->file_source_ = FileSource::EXISTING_FILE;
            
            // 🎯 API修复：为保存功能准备FileManager
            if (!loaded_workbook->open()) {
                FASTEXCEL_LOG_ERROR("Failed to prepare FileManager for workbook: {}", path.string());
                return nullptr;
            }
        }
        
        FASTEXCEL_LOG_INFO("Successfully loaded workbook for editing: {}", path.string());
        return loaded_workbook;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Exception while loading workbook for editing: {}, error: {}", path.string(), e.what());
        return nullptr;
    }
}

std::unique_ptr<Workbook> Workbook::open(const std::string& filepath) {
    return open(Path(filepath));
}

bool Workbook::refresh() {
    
    try {
        // 保存当前状态
        std::string current_filename = filename_;
        
        // 关闭当前工作簿
        close();
        
        // 重新加载
        Path current_path(current_filename);
        auto refreshed_workbook = open(current_path);
        if (!refreshed_workbook) {
            FASTEXCEL_LOG_ERROR("Failed to refresh workbook: {}", current_filename);
            return false;
        }
        
        // 替换当前内容
        worksheets_ = std::move(refreshed_workbook->worksheets_);
        format_repo_ = std::move(refreshed_workbook->format_repo_);
        
        // 通过管理器复制文档属性
        if (refreshed_workbook->document_manager_ && document_manager_) {
            document_manager_->setDocumentProperties(refreshed_workbook->document_manager_->getDocumentProperties());
            // 复制自定义属性
            auto custom_props = refreshed_workbook->document_manager_->getAllCustomProperties();
            for (const auto& [name, value] : custom_props) {
                document_manager_->setCustomProperty(name, value);
            }
        }
        
        // 重新打开工作簿
        open();
        
        FASTEXCEL_LOG_INFO("Successfully refreshed workbook: {}", current_filename);
        return true;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Exception during workbook refresh: {}", e.what());
        return false;
    }
}

bool Workbook::mergeWorkbook(const std::unique_ptr<Workbook>& other_workbook) {
    return mergeWorkbook(other_workbook, MergeOptions{});
}

bool Workbook::mergeWorkbook(const std::unique_ptr<Workbook>& other_workbook, const MergeOptions& options) {
    if (!other_workbook) {
        FASTEXCEL_LOG_ERROR("Cannot merge: other workbook is null");
        return false;
    }
    
    
    try {
        int merged_count = 0;
        
        // 合并工作表
        if (options.merge_worksheets) {
            for (const auto& other_worksheet : other_workbook->worksheets_) {
                std::string new_name = options.name_prefix + other_worksheet->getName();
                
                // 检查名称冲突
                if (getSheet(new_name) != nullptr) {
                    if (options.overwrite_existing) {
                        removeSheet(new_name);
                        FASTEXCEL_LOG_INFO("Removed existing worksheet for merge: {}", new_name);
                    } else {
                        new_name = generateUniqueSheetName(new_name);
                        FASTEXCEL_LOG_INFO("Generated unique name for merge: {}", new_name);
                    }
                }
                
                // 创建新工作表并复制内容
                auto new_worksheet = addSheet(new_name);
                if (new_worksheet) {
                    // 实现深拷贝逻辑：复制所有单元格、格式、设置等
                    try {
                        auto [max_row, max_col] = other_worksheet->getUsedRange();
                        
                        // 复制所有单元格和格式
                        for (int row = 0; row <= max_row; ++row) {
                            for (int col = 0; col <= max_col; ++col) {
                                if (other_worksheet->hasCellAt(row, col)) {
                                    const auto& source_cell = other_worksheet->getCell(row, col);
                                    auto& new_cell = new_worksheet->getCell(row, col);
                                    
                                    // 复制单元格值
                                    new_cell = source_cell;
                                    
                                    // 复制格式
                                    auto format = source_cell.getFormatDescriptor();
                                    if (format) {
                                        new_worksheet->setCellFormat(row, col, format);
                                    }
                                }
                            }
                        }
                        
                        // 复制工作表设置
                        new_worksheet->setTabSelected(other_worksheet->isTabSelected());
                        
                        // 复制列宽和行高
                        for (int col = 0; col <= max_col; ++col) {
                            auto width_opt = other_worksheet->tryGetColumnWidth(col);
                            if (width_opt && *width_opt > 0) {
                                new_worksheet->setColumnWidth(col, *width_opt);
                            }
                        }
                        
                        for (int row = 0; row <= max_row; ++row) {
                            auto height_opt = other_worksheet->tryGetRowHeight(row);
                            if (height_opt && *height_opt > 0) {
                                new_worksheet->setRowHeight(row, *height_opt);
                            }
                        }
                        
                        FASTEXCEL_LOG_DEBUG("Deep copied {} cells from {} to {}", 
                                           (max_row + 1) * (max_col + 1), other_worksheet->getName(), new_name);
                    } catch (const std::exception& e) {
                        FASTEXCEL_LOG_ERROR("Failed to copy worksheet content during merge: {}", e.what());
                    }
                    
                    merged_count++;
                    FASTEXCEL_LOG_DEBUG("Merged worksheet: {} -> {}", other_worksheet->getName(), new_name);
                }
            }
        }
        
        // 合并格式
        if (options.merge_formats) {
            // 将其他工作簿的格式仓储合并到当前格式仓储
            // 使用线程安全的快照方式遍历其他工作簿的所有格式并添加到当前仓储中（自动去重）
            auto format_snapshot = other_workbook->format_repo_->createSnapshot();
            for (const auto& format_item : format_snapshot) {
                format_repo_->addFormat(*format_item.second);
            }
            FASTEXCEL_LOG_DEBUG("Merged formats from other workbook");
        }
        
        // 合并文档属性
        if (options.merge_properties && other_workbook->document_manager_ && document_manager_) {
            const auto& other_props = other_workbook->document_manager_->getDocumentProperties();
            if (!other_props.title.empty()) {
                document_manager_->setTitle(other_props.title);
            }
            if (!other_props.author.empty()) {
                document_manager_->setAuthor(other_props.author);
            }
            if (!other_props.subject.empty()) {
                document_manager_->setSubject(other_props.subject);
            }
            if (!other_props.company.empty()) {
                document_manager_->setCompany(other_props.company);
            }
            
            // 合并自定义属性
            auto custom_props = other_workbook->document_manager_->getAllCustomProperties();
            for (const auto& [name, value] : custom_props) {
                setProperty(name, value);
            }
            
            FASTEXCEL_LOG_DEBUG("Merged document properties");
        }
        
        FASTEXCEL_LOG_INFO("Successfully merged workbook: {} worksheets, {} formats",
                merged_count, other_workbook->format_repo_->getFormatCount());
        return true;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Exception during workbook merge: {}", e.what());
        return false;
    }
}

bool Workbook::exportWorksheets(const std::vector<std::string>& worksheet_names, const std::string& output_filename) {
    if (worksheet_names.empty()) {
        FASTEXCEL_LOG_ERROR("No worksheets specified for export");
        return false;
    }
    
    try {
        // 创建新工作簿
        auto export_workbook = create(Path(output_filename));
        if (!export_workbook->open()) {
            FASTEXCEL_LOG_ERROR("Failed to create export workbook: {}", output_filename);
            return false;
        }
        
        // 复制指定的工作表
        int exported_count = 0;
        for (const std::string& name : worksheet_names) {
            auto source_worksheet = getSheet(name);
            if (!source_worksheet) {
                FASTEXCEL_LOG_WARN("Worksheet not found for export: {}", name);
                continue;
            }
            
            auto new_worksheet = export_workbook->addSheet(name);
            if (new_worksheet) {
                // 这里需要实现深拷贝逻辑
                // 简化版本：复制基本属性
                exported_count++;
                FASTEXCEL_LOG_DEBUG("Exported worksheet: {}", name);
            }
        }
        
        // 复制文档属性
        if (document_manager_ && export_workbook->document_manager_) {
            export_workbook->document_manager_->setDocumentProperties(document_manager_->getDocumentProperties());
            // 复制自定义属性
            auto custom_props = document_manager_->getAllCustomProperties();
            for (const auto& [name, value] : custom_props) {
                export_workbook->setProperty(name, value);
            }
        }
        
        // 保存导出的工作簿
        bool success = export_workbook->save();
        export_workbook->close();
        
        if (success) {
            FASTEXCEL_LOG_INFO("Successfully exported {} worksheets to: {}", exported_count, output_filename);
        } else {
            FASTEXCEL_LOG_ERROR("Failed to save exported workbook: {}", output_filename);
        }
        
        return success;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Exception during worksheet export: {}", e.what());
        return false;
    }
}

int Workbook::batchRenameWorksheets(const std::unordered_map<std::string, std::string>& rename_map) {
    int renamed_count = 0;
    
    for (const auto& [old_name, new_name] : rename_map) {
        if (renameSheet(old_name, new_name)) {
            renamed_count++;
            FASTEXCEL_LOG_DEBUG("Renamed worksheet: {} -> {}", old_name, new_name);
        } else {
            FASTEXCEL_LOG_WARN("Failed to rename worksheet: {} -> {}", old_name, new_name);
        }
    }
    
    FASTEXCEL_LOG_INFO("Batch rename completed: {} worksheets renamed", renamed_count);
    return renamed_count;
}

int Workbook::batchRemoveWorksheets(const std::vector<std::string>& worksheet_names) {
    int removed_count = 0;
    
    for (const std::string& name : worksheet_names) {
        if (removeSheet(name)) {
            removed_count++;
            FASTEXCEL_LOG_DEBUG("Removed worksheet: {}", name);
        } else {
            FASTEXCEL_LOG_WARN("Failed to remove worksheet: {}", name);
        }
    }
    
    FASTEXCEL_LOG_INFO("Batch remove completed: {} worksheets removed", removed_count);
    return removed_count;
}

bool Workbook::reorderWorksheets(const std::vector<std::string>& new_order) {
    if (new_order.size() != worksheets_.size()) {
        FASTEXCEL_LOG_ERROR("New order size ({}) doesn't match worksheet count ({})",
                 new_order.size(), worksheets_.size());
        return false;
    }
    
    try {
        std::vector<std::shared_ptr<Worksheet>> reordered_worksheets;
        reordered_worksheets.reserve(worksheets_.size());
        
        // 按新顺序重新排列工作表
        for (const std::string& name : new_order) {
            auto worksheet = getSheet(name);
            if (!worksheet) {
                FASTEXCEL_LOG_ERROR("Worksheet not found in reorder list: {}", name);
                return false;
            }
            reordered_worksheets.push_back(worksheet);
        }
        
        // 替换工作表列表
        worksheets_ = std::move(reordered_worksheets);
        
        FASTEXCEL_LOG_INFO("Successfully reordered {} worksheets", worksheets_.size());
        return true;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Exception during worksheet reordering: {}", e.what());
        return false;
    }
}

int Workbook::findAndReplaceAll(const std::string& find_text, const std::string& replace_text,
                         const FindReplaceOptions& options) {
    int total_replacements = 0;
    
    for (const auto& worksheet : worksheets_) {
        // 检查工作表过滤器
        if (!options.worksheet_filter.empty()) {
            bool found = std::find(options.worksheet_filter.begin(), options.worksheet_filter.end(),
                                 worksheet->getName()) != options.worksheet_filter.end();
            if (!found) {
                continue; // 跳过不在过滤器中的工作表
            }
        }
        
        int replacements = worksheet->findAndReplace(find_text, replace_text,
                                                   options.match_case, options.match_entire_cell);
        total_replacements += replacements;
        
        if (replacements > 0) {
            FASTEXCEL_LOG_DEBUG("Found and replaced {} occurrences in worksheet: {}",
                     replacements, worksheet->getName());
        }
    }
    
    FASTEXCEL_LOG_INFO("Global find and replace completed: {} total replacements", total_replacements);
    return total_replacements;
}

int Workbook::findAndReplaceAll(const std::string& find_text, const std::string& replace_text) {
    return findAndReplaceAll(find_text, replace_text, FindReplaceOptions{});
}

std::vector<std::tuple<std::string, int, int>> Workbook::findAll(const std::string& search_text,
                                                           const FindReplaceOptions& options) {
    std::vector<std::tuple<std::string, int, int>> results;
    
    for (const auto& worksheet : worksheets_) {
        // 检查工作表过滤器
        if (!options.worksheet_filter.empty()) {
            bool found = std::find(options.worksheet_filter.begin(), options.worksheet_filter.end(),
                                 worksheet->getName()) != options.worksheet_filter.end();
            if (!found) {
                continue; // 跳过不在过滤器中的工作表
            }
        }
        
        auto worksheet_results = worksheet->findCells(search_text, options.match_case, options.match_entire_cell);
        
        // 将结果添加到总结果中，包含工作表名称
        for (const auto& [row, col] : worksheet_results) {
            results.emplace_back(worksheet->getName(), row, col);
        }
        
        if (!worksheet_results.empty()) {
            FASTEXCEL_LOG_DEBUG("Found {} matches in worksheet: {}", worksheet_results.size(), worksheet->getName());
        }
    }
    
    FASTEXCEL_LOG_INFO("Global search completed: {} total matches found", results.size());
    return results;
}

std::vector<std::tuple<std::string, int, int>> Workbook::findAll(const std::string& search_text) {
    return findAll(search_text, FindReplaceOptions{});
}

Workbook::WorkbookStats Workbook::getStatistics() const {
    WorkbookStats stats;
    
    stats.total_worksheets = worksheets_.size();
    stats.total_formats = format_repo_->getFormatCount();
    
    // 计算总单元格数和内存使用
    for (const auto& worksheet : worksheets_) {
        size_t cell_count = worksheet->getCellCount();
        stats.total_cells += cell_count;
        stats.worksheet_cell_counts[worksheet->getName()] = cell_count;
        
        if (worksheet->isOptimizeMode()) {
            stats.memory_usage += worksheet->getMemoryUsage();
        }
    }
    
    // 估算工作簿本身的内存使用
    stats.memory_usage += sizeof(Workbook);
    stats.memory_usage += worksheets_.capacity() * sizeof(std::shared_ptr<Worksheet>);
    stats.memory_usage += format_repo_->getMemoryUsage();
    
    // 估算管理器的内存使用量
    if (document_manager_) {
        stats.memory_usage += document_manager_->getCustomPropertyCount() * 64; // 估算每个属性64字节
    }
    
    return stats;
}

// 智能模式选择辅助方法

size_t Workbook::estimateMemoryUsage() const {
    size_t total_memory = 0;
    
    // 估算工作表内存使用
    for (const auto& worksheet : worksheets_) {
        if (worksheet->isOptimizeMode()) {
            total_memory += worksheet->getMemoryUsage();
        } else {
            // 估算标准模式的内存使用
            auto [max_row, max_col] = worksheet->getUsedRange();
            if (max_row >= 0 && max_col >= 0) {
                size_t cell_count = (max_row + 1) * (max_col + 1);
                total_memory += cell_count * 100; // 估算每个单元格100字节
            }
        }
    }
    
    // 估算格式池内存
    total_memory += format_repo_->getMemoryUsage();
    
    // 估算共享字符串内存
    if (shared_string_table_) {
        total_memory += shared_string_table_->getMemoryUsage();
    }
    
    // 估算XML生成时的临时内存（约为数据的2-3倍）
    total_memory *= 3;
    
    return total_memory;
}

size_t Workbook::getTotalCellCount() const {
    size_t total_cells = 0;
    
    for (const auto& worksheet : worksheets_) {
        if (worksheet->isOptimizeMode()) {
            total_cells += worksheet->getCellCount();
        } else {
            auto [max_row, max_col] = worksheet->getUsedRange();
            if (max_row >= 0 && max_col >= 0) {
                // 估算实际有数据的单元格数量（不是整个矩形区域）
                size_t estimated_cells = 0;
                for (int row = 0; row <= max_row; ++row) {
                    for (int col = 0; col <= max_col; ++col) {
                        if (worksheet->hasCellAt(row, col)) {
                            estimated_cells++;
                        }
                    }
                }
                total_cells += estimated_cells;
            }
        }
    }
    
    return total_cells;
}

size_t Workbook::getEstimatedSize() const {
    // 估计文件大小：基础大小 + 工作表大小 + 样式大小 + 共享字符串大小
    size_t estimated = 10 * 1024; // 基础XML文件大約10KB
    
    // 每个工作表的估计大小
    for (const auto& sheet : worksheets_) {
        if (sheet) {
            // 每个单元格平均约50字节（XML格式）
            estimated += sheet->getCellCount() * 50;
            // 每个工作表的基础结构约5KB
            estimated += 5 * 1024;
        }
    }
    
    // 样式大小估计
    if (format_repo_) {
        estimated += format_repo_->getFormatCount() * 200; // 每个样式约200字节
    }
    
    // 共享字符串大小估计
    if (shared_string_table_) {
        estimated += shared_string_table_->getStringCount() * 30; // 每个字符串平均30字节
    }
    
    return estimated;
}

std::unique_ptr<StyleTransferContext> Workbook::copyStylesFrom(const Workbook& source_workbook) {
    FASTEXCEL_LOG_DEBUG("开始从源工作簿复制样式数据");
    
    // 创建样式传输上下文
    auto transfer_context = std::make_unique<StyleTransferContext>(*source_workbook.format_repo_, *format_repo_);
    
    // 预加载所有映射以触发批量复制
    transfer_context->preloadAllMappings();
    
    auto stats = transfer_context->getTransferStats();
    FASTEXCEL_LOG_DEBUG("完成样式复制，传输了{}个格式，去重了{}个", 
             stats.transferred_count, stats.deduplicated_count);
    
    // 自动复制主题 XML 以保持颜色和字体一致性
    const std::string& source_theme = source_workbook.getThemeXML();
    if (!source_theme.empty()) {
        // 只有当前工作簿没有自定义主题时才复制源主题
        if (theme_xml_.empty()) {
            theme_xml_ = source_theme;
            FASTEXCEL_LOG_DEBUG("自动复制主题XML ({} 字节)", theme_xml_.size());
        } else {
            FASTEXCEL_LOG_DEBUG("当前工作簿已有自定义主题，保持现有主题不变");
        }
    } else {
        FASTEXCEL_LOG_DEBUG("源工作簿无自定义主题，保持默认主题");
    }
    
    return transfer_context;
}

FormatRepository::DeduplicationStats Workbook::getStyleStats() const {
    return format_repo_->getDeduplicationStats();
}

bool Workbook::generateWithGenerator(bool use_streaming_writer) {
    if (!file_manager_) {
        FASTEXCEL_LOG_ERROR("FileManager is null - cannot write workbook");
        return false;
    }
    std::unique_ptr<IFileWriter> writer;
    if (use_streaming_writer) {
        writer = std::make_unique<StreamingFileWriter>(file_manager_.get());
    } else {
        writer = std::make_unique<BatchFileWriter>(file_manager_.get());
    }
    ExcelStructureGenerator generator(this, std::move(writer));
    return generator.generate();
}

bool Workbook::isModified() const {
    // 检查DirtyManager是否有修改标记
    auto* dirty_manager = getDirtyManager();
    if (dirty_manager && dirty_manager->hasDirtyData()) {
        return true;
    }
    
    // 检查主题是否被修改
    if (theme_dirty_) {
        return true;
    }
    
    // 检查工作表是否有修改（如果有hasChanges方法）
    for (const auto& worksheet : worksheets_) {
        if (worksheet) {
            // TODO: 检查Worksheet是否有hasChanges或类似方法
            // if (worksheet->hasChanges()) {
            //     return true;
            // }
        }
    }
    
    return false;
}

// 访问模式检查辅助方法实现

void Workbook::ensureEditable(const std::string& operation) const {
    if (state_ == WorkbookState::READING) {
        std::string msg = "Cannot perform operation";
        if (!operation.empty()) {
            msg += " '" + operation + "'";
        }
        msg += ": workbook is opened in read-only mode. Use openForEditing() instead of openForReading().";
        
        FASTEXCEL_LOG_ERROR("{}", msg);
        throw OperationException(msg, operation);
    }
}

void Workbook::ensureReadable(const std::string& operation) const {
    // 读取操作在任何模式下都是允许的
    // 这个方法预留用于未来可能的扩展，比如检查文件是否损坏等
    (void)operation; // 避免未使用参数警告
}

// 新状态管理系统实现

bool Workbook::isStateValid(WorkbookState required_state) const {
    // 状态层级：CLOSED < CREATING/READING/EDITING
    // CREATING/READING/EDITING 是平级的，但有不同的权限
    
    switch (required_state) {
        case WorkbookState::CLOSED:
            return true; // 任何状态都可以关闭
            
        case WorkbookState::CREATING:
            return state_ == WorkbookState::CREATING;
            
        case WorkbookState::READING:
            // 读取操作在 READING 和 EDITING 状态都允许
            return state_ == WorkbookState::READING || state_ == WorkbookState::EDITING;
            
        case WorkbookState::EDITING:
            // 编辑操作只在 EDITING 和 CREATING 状态允许
            return state_ == WorkbookState::EDITING || state_ == WorkbookState::CREATING;
            
        default:
            return false;
    }
}

void Workbook::transitionToState(WorkbookState new_state, const std::string& reason) {
    if (state_ == new_state) {
        return; // 状态未改变
    }
    
    WorkbookState old_state = state_;
    state_ = new_state;
    
    FASTEXCEL_LOG_DEBUG("Workbook state transition: {} -> {} ({})", 
              static_cast<int>(old_state), 
              static_cast<int>(new_state), 
              reason.empty() ? "no reason" : reason);
}

// 样式构建器

StyleBuilder Workbook::createStyleBuilder() const {
    return StyleBuilder();
}

}} // namespace fastexcel::core
